#!/usr/bin/env python3
"""Import XLSX, CSV, and TSV spreadsheets into SQLite tables."""

from __future__ import annotations

import argparse
import base64
from contextlib import contextmanager
import csv
import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import time
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Iterable, List, Optional, Union

from openpyxl import load_workbook


PROJECT_DIR = Path(__file__).resolve().parent
SPREADSHEET_EXTENSIONS = {".xlsx", ".csv", ".tsv"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".heic", ".heif"}
SUPPORTED_EXTENSIONS = SPREADSHEET_EXTENSIONS | IMAGE_EXTENSIONS


def load_env_file(path: Path) -> None:
    """Load simple KEY=VALUE settings without an extra dependency."""
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


load_env_file(PROJECT_DIR / ".env")


def project_path(value: str) -> Path:
    path = Path(value).expanduser()
    return path.resolve() if path.is_absolute() else (PROJECT_DIR / path).resolve()


@dataclass(frozen=True)
class Settings:
    input_dir: Path = project_path(os.getenv("INPUT_DIR", "incoming"))
    processed_dir: Path = project_path(os.getenv("PROCESSED_DIR", "processed"))
    failed_dir: Path = project_path(os.getenv("FAILED_DIR", "failed"))
    database_path: Path = project_path(os.getenv("DATABASE_PATH", "data/imports.db"))
    add_source_columns: bool = os.getenv("ADD_SOURCE_COLUMNS", "true").lower() != "false"
    watch_interval_seconds: float = float(os.getenv("WATCH_INTERVAL_SECONDS", "2"))
    tesseract_command: str = os.getenv("TESSERACT_COMMAND", "tesseract")
    ocr_language: str = os.getenv("OCR_LANGUAGE", "eng")
    ocr_min_confidence: float = float(os.getenv("OCR_MIN_CONFIDENCE", "20"))
    ocr_acceptance_confidence: float = float(os.getenv("OCR_ACCEPTANCE_CONFIDENCE", "60"))
    image_ocr_engine: str = os.getenv("IMAGE_OCR_ENGINE", "openai").lower()
    openai_model: str = os.getenv("OPENAI_MODEL", "gpt-6-luna")


SETTINGS = Settings()


def quote_identifier(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def clean_identifier(value: Any, fallback: str = "column") -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(character for character in text if not unicodedata.combining(character))
    text = re.sub(r"[^a-z0-9_]+", "_", text.strip().lower())
    text = re.sub(r"_+", "_", text).strip("_") or fallback
    if text[0].isdigit():
        text = f"column_{text}"
    return text[:63]


def unique_headers(header_row: list[Any], width: int) -> list[str]:
    counts: dict[str, int] = {}
    headers: list[str] = []
    for index in range(width):
        base = clean_identifier(header_row[index] if index < len(header_row) else None, f"column_{index + 1}")
        counts[base] = counts.get(base, 0) + 1
        headers.append(base if counts[base] == 1 else f"{base}_{counts[base]}")
    return headers


def is_empty(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def infer_type(values: Iterable[Any]) -> str:
    populated = [value for value in values if not is_empty(value)]
    if not populated:
        return "TEXT"
    if all(isinstance(value, (bool, int)) and not isinstance(value, float) for value in populated):
        return "INTEGER"
    if all(isinstance(value, (bool, int, float)) for value in populated):
        return "REAL"
    return "TEXT"


def normalize_value(value: Any, sqlite_type: str) -> Any:
    if is_empty(value):
        return None
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, bool):
        return int(value)
    if sqlite_type == "INTEGER" and isinstance(value, (int, float)):
        return int(value)
    if sqlite_type == "REAL" and isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, (dict, list, tuple)):
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def calculate_file_hash(file_path: Path) -> str:
    digest = hashlib.sha256()
    with file_path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def initialize_database(database: sqlite3.Connection) -> None:
    database.executescript("""
        PRAGMA journal_mode = WAL;
        CREATE TABLE IF NOT EXISTS _spreadsheet_import_log (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          source_file TEXT NOT NULL,
          source_sheet TEXT NOT NULL,
          file_hash TEXT NOT NULL,
          table_name TEXT,
          rows_imported INTEGER NOT NULL DEFAULT 0,
          status TEXT NOT NULL CHECK (status IN ('success', 'skipped', 'failed')),
          message TEXT,
          imported_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_import_log_lookup
          ON _spreadsheet_import_log(file_hash, source_sheet, table_name, status);
    """)


def log_import(database: sqlite3.Connection, source_file: str, source_sheet: str,
               file_hash: str, table_name: str | None, status: str,
               rows_imported: int = 0, message: str | None = None) -> None:
    database.execute("""
        INSERT INTO _spreadsheet_import_log
          (source_file, source_sheet, file_hash, table_name, rows_imported, status, message, imported_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (source_file, source_sheet, file_hash, table_name, rows_imported, status, message, now_iso()))


def ensure_table(database: sqlite3.Connection, table_name: str,
                 columns: list[tuple[str, str]]) -> None:
    exists = database.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?", (table_name,)
    ).fetchone()
    additions = list(columns)
    if SETTINGS.add_source_columns:
        additions.extend([
            ("_source_file", "TEXT"), ("_source_sheet", "TEXT"), ("_imported_at", "TEXT")
        ])
    if not exists:
        definitions = ", ".join(
            f"{quote_identifier(name)} {sqlite_type}" for name, sqlite_type in additions
        )
        database.execute(f"CREATE TABLE {quote_identifier(table_name)} ({definitions})")
        return
    existing = {
        row[1] for row in database.execute(f"PRAGMA table_info({quote_identifier(table_name)})").fetchall()
    }
    for name, sqlite_type in additions:
        if name not in existing:
            database.execute(
                f"ALTER TABLE {quote_identifier(table_name)} ADD COLUMN {quote_identifier(name)} {sqlite_type}"
            )


def parse_matrix(matrix: list[list[Any]]) -> tuple[list[tuple[str, str]], list[list[Any]]] | None:
    while matrix and all(is_empty(value) for value in matrix[-1]):
        matrix.pop()
    if len(matrix) < 2:
        return None
    width = max(len(row) for row in matrix)
    headers = unique_headers(matrix[0], width)
    rows = [
        list(row) + [None] * (width - len(row))
        for row in matrix[1:] if not all(is_empty(value) for value in row)
    ]
    if not rows:
        return None
    columns = [(name, infer_type(row[index] for row in rows)) for index, name in enumerate(headers)]
    return columns, rows


def read_xlsx(file_path: Path) -> list[tuple[str, list[list[Any]]]]:
    workbook = load_workbook(file_path, read_only=True, data_only=True)
    try:
        return [
            (worksheet.title, [list(row) for row in worksheet.iter_rows(values_only=True)])
            for worksheet in workbook.worksheets
        ]
    finally:
        workbook.close()


def read_delimited(file_path: Path) -> list[tuple[str, list[list[Any]]]]:
    delimiter = "\t" if file_path.suffix.lower() == ".tsv" else ","
    with file_path.open("r", encoding="utf-8-sig", newline="") as source:
        return [("Data", [list(row) for row in csv.reader(source, delimiter=delimiter)])]


@contextmanager
def normalized_image_path(file_path: Path):
    """Yield a directly readable image path, converting HEIC/HEIF to temporary JPEG."""
    if file_path.suffix.lower() not in {".heic", ".heif"}:
        yield file_path
        return
    try:
        from PIL import Image
        from pillow_heif import register_heif_opener
    except ImportError as error:
        raise RuntimeError("Install HEIC support with: pip install -r requirements.txt") from error

    register_heif_opener()
    temporary_path: Path | None = None
    try:
        with Image.open(file_path) as image:
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as temporary:
                temporary_path = Path(temporary.name)
            image.convert("RGB").save(temporary_path, format="JPEG", quality=95)
        yield temporary_path
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def clean_ocr_text(text: str) -> str:
    text = text.strip().strip("|[]{}")
    return re.sub(r"\s+", " ", text).strip()


def coerce_ocr_value(value: str) -> Any:
    text = clean_ocr_text(value)
    if not text or text in {"-", "—", "_"}:
        return None
    number = text.replace(",", ".")
    if re.fullmatch(r"[-+]?\d+", number):
        return int(number)
    if re.fullmatch(r"[-+]?(?:\d+\.\d+|\.\d+)", number):
        return float(number)
    return text


def group_ocr_rows(words: list[dict[str, Any]], tolerance: float) -> list[list[dict[str, Any]]]:
    rows: list[list[dict[str, Any]]] = []
    for word in sorted(words, key=lambda item: (item["center_y"], item["left"])):
        matching = next(
            (row for row in rows if abs(sum(item["center_y"] for item in row) / len(row) - word["center_y"]) <= tolerance),
            None,
        )
        if matching is None:
            rows.append([word])
        else:
            matching.append(word)
    return [sorted(row, key=lambda item: item["left"]) for row in rows]


def header_groups(header_words: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    heights = sorted(word["height"] for word in header_words)
    typical_height = heights[len(heights) // 2] if heights else 25
    gap_threshold = max(45, typical_height * 1.35)
    groups: list[list[dict[str, Any]]] = []
    for word in sorted(header_words, key=lambda item: item["left"]):
        if not groups:
            groups.append([word])
            continue
        previous = groups[-1][-1]
        gap = word["left"] - (previous["left"] + previous["width"])
        if gap > gap_threshold:
            groups.append([word])
        else:
            groups[-1].append(word)
    return groups


def read_image_with_tesseract(file_path: Path) -> list[tuple[str, list[list[Any]]]]:
    """Use local Tesseract OCR to convert a photographed table to a matrix."""
    try:
        process = subprocess.run(
            [SETTINGS.tesseract_command, str(file_path), "stdout", "--psm", "6", "-l", SETTINGS.ocr_language, "tsv"],
            check=True,
            capture_output=True,
            text=True,
        )
    except FileNotFoundError as error:
        raise RuntimeError(
            "Tesseract OCR is not installed. On macOS run: brew install tesseract"
        ) from error
    except subprocess.CalledProcessError as error:
        raise RuntimeError(f"Tesseract could not read the image: {error.stderr.strip()}") from error

    records = list(csv.DictReader(process.stdout.splitlines(), delimiter="\t"))
    words: list[dict[str, Any]] = []
    for record in records:
        text = clean_ocr_text(record.get("text", ""))
        try:
            confidence = float(record.get("conf", "-1"))
        except ValueError:
            confidence = -1
        if not text or confidence < SETTINGS.ocr_min_confidence:
            continue
        if not re.search(r"[A-Za-z0-9#]", text):
            continue
        left = int(record["left"])
        top = int(record["top"])
        width = int(record["width"])
        height = int(record["height"])
        words.append({
            "text": text,
            "confidence": confidence,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
            "center_x": left + width / 2,
            "center_y": top + height / 2,
        })
    if not words:
        raise ValueError("No readable text was found in the image.")

    heights = sorted(word["height"] for word in words)
    typical_height = heights[len(heights) // 2]
    rows = group_ocr_rows(words, max(16, typical_height * 0.75))
    header_keywords = {"user", "id", "shift", "pattern", "uph", "hrs", "audits", "auditor"}
    header_row = max(
        rows,
        key=lambda row: sum(
            1 for word in row
            if re.sub(r"[^a-z]", "", word["text"].lower()) in header_keywords
        ),
    )
    keyword_score = sum(
        1 for word in header_row
        if re.sub(r"[^a-z]", "", word["text"].lower()) in header_keywords
    )
    if keyword_score < 2:
        raise ValueError(
            "Could not identify the table header. Crop the photo to the table and try again."
        )

    groups = header_groups(header_row)
    if len(groups) < 2:
        raise ValueError("OCR found fewer than two table columns.")
    headers = [" ".join(word["text"] for word in group) for group in groups]
    centers = [sum(word["center_x"] for word in group) / len(group) for group in groups]
    boundaries = [(centers[index] + centers[index + 1]) / 2 for index in range(len(centers) - 1)]
    header_bottom = max(word["top"] + word["height"] for word in header_row)

    matrix: list[list[Any]] = [headers + ["OCR Confidence", "Review Required"]]
    row_confidences: list[float] = []
    for row in rows:
        if sum(word["center_y"] for word in row) / len(row) <= header_bottom:
            continue
        cells: list[list[dict[str, Any]]] = [[] for _ in headers]
        for word in row:
            column = sum(1 for boundary in boundaries if word["center_x"] > boundary)
            if 0 <= column < len(cells):
                cells[column].append(word)
        values = [
            coerce_ocr_value(" ".join(word["text"] for word in sorted(cell, key=lambda item: item["left"])))
            for cell in cells
        ]
        if sum(value is not None for value in values) < 2:
            continue
        confidence = round(sum(word["confidence"] for word in row) / len(row), 1)
        row_confidences.append(confidence)
        matrix.append(values + [confidence, True])
    if len(matrix) < 2:
        raise ValueError("The table header was found, but no data rows could be extracted.")
    overall_confidence = sum(row_confidences) / len(row_confidences)
    if overall_confidence < SETTINGS.ocr_acceptance_confidence:
        raise ValueError(
            f"Local OCR confidence was only {overall_confidence:.1f}%. "
            "The image was not imported because the results could be unreliable. "
            "Use IMAGE_OCR_ENGINE=openai for difficult photos, or provide a clearer cropped image."
        )
    return [("OCR Data", matrix)]


def read_image_with_openai(file_path: Path) -> list[tuple[str, list[list[Any]]]]:
    """Extract a photographed table with an image-capable OpenAI model."""
    if not os.getenv("OPENAI_API_KEY"):
        raise RuntimeError(
            "Photo import needs OPENAI_API_KEY when IMAGE_OCR_ENGINE=openai. "
            "Spreadsheet, CSV, and TSV imports do not need an API key."
        )
    try:
        from openai import OpenAI
        from pydantic import BaseModel
    except ImportError as error:
        raise RuntimeError("Install photo-import dependencies with: pip install -r requirements.txt") from error

    class ExtractedTable(BaseModel):
        title: Optional[str]
        headers: List[str]
        rows: List[List[Union[str, int, float, bool, None]]]
        notes: List[str]

    mime_type = "image/png" if file_path.suffix.lower() == ".png" else "image/jpeg"
    image_data = base64.b64encode(file_path.read_bytes()).decode("ascii")
    client = OpenAI()
    response = client.responses.parse(
        model=SETTINGS.openai_model,
        input=[{
            "role": "user",
            "content": [
                {
                    "type": "input_text",
                    "text": (
                        "Extract the visible spreadsheet table exactly. Preserve every column and data row. "
                        "Do not include Excel row numbers, column letters, tabs, or the title as data. "
                        "Keep identifiers as strings. Keep numeric values as numbers. Use null when a cell is "
                        "blank or genuinely unreadable; never guess an unreadable value. Return rows in visual order."
                    ),
                },
                {
                    "type": "input_image",
                    "image_url": f"data:{mime_type};base64,{image_data}",
                    "detail": "original",
                },
            ],
        }],
        text_format=ExtractedTable,
    )
    extracted = response.output_parsed
    if extracted is None or len(extracted.headers) < 2 or not extracted.rows:
        raise ValueError("The model could not find a usable table in the image.")
    width = len(extracted.headers)
    matrix: list[list[Any]] = [extracted.headers + ["Review Required"]]
    for source_row in extracted.rows:
        row = list(source_row[:width]) + [None] * max(0, width - len(source_row))
        matrix.append(row + [True])
    if extracted.notes:
        print("Photo extraction notes: " + "; ".join(extracted.notes))
    return [("OCR Data", matrix)]


def read_image_table(file_path: Path) -> list[tuple[str, list[list[Any]]]]:
    with normalized_image_path(file_path) as readable_path:
        if SETTINGS.image_ocr_engine == "openai":
            return read_image_with_openai(readable_path)
        if SETTINGS.image_ocr_engine == "tesseract":
            return read_image_with_tesseract(readable_path)
        raise ValueError("IMAGE_OCR_ENGINE must be 'openai' or 'tesseract'.")


def read_spreadsheet(file_path: Path) -> list[tuple[str, list[list[Any]]]]:
    extension = file_path.suffix.lower()
    if extension in IMAGE_EXTENSIONS:
        return read_image_table(file_path)
    return read_xlsx(file_path) if extension == ".xlsx" else read_delimited(file_path)


def destination_table(file_path: Path, sheet_name: str, sheet_count: int,
                      override: str | None) -> str:
    if override:
        requested = override if sheet_count == 1 else f"{override}_{sheet_name}"
        return clean_identifier(requested, "imported_data")
    return clean_identifier(f"{file_path.stem}_{sheet_name}", "imported_data")


def import_sheet(database: sqlite3.Connection, file_path: Path, file_hash: str,
                 sheet_name: str, matrix: list[list[Any]], sheet_count: int,
                 table_override: str | None) -> int:
    table_name = destination_table(file_path, sheet_name, sheet_count, table_override)
    source_file = file_path.name
    already_imported = database.execute("""
        SELECT 1 FROM _spreadsheet_import_log
        WHERE file_hash = ? AND source_sheet = ? AND table_name = ? AND status = 'success'
        LIMIT 1
    """, (file_hash, sheet_name, table_name)).fetchone()
    if already_imported:
        log_import(database, source_file, sheet_name, file_hash, table_name, "skipped",
                   message="This file and sheet were already imported.")
        database.commit()
        print(f"Skipped {source_file} / {sheet_name}: already imported.")
        return 0
    parsed = parse_matrix(matrix)
    if not parsed:
        log_import(database, source_file, sheet_name, file_hash, table_name, "skipped",
                   message="No data rows found.")
        database.commit()
        print(f"Skipped {source_file} / {sheet_name}: no data rows.")
        return 0
    columns, rows = parsed
    try:
        database.execute("BEGIN")
        ensure_table(database, table_name, columns)
        insert_columns = [name for name, _ in columns]
        if SETTINGS.add_source_columns:
            insert_columns.extend(["_source_file", "_source_sheet", "_imported_at"])
        placeholders = ", ".join("?" for _ in insert_columns)
        sql = (f"INSERT INTO {quote_identifier(table_name)} "
               f"({', '.join(quote_identifier(name) for name in insert_columns)}) VALUES ({placeholders})")
        imported_at = now_iso()
        prepared_rows = []
        for row in rows:
            values = [normalize_value(row[index], sqlite_type)
                      for index, (_, sqlite_type) in enumerate(columns)]
            if SETTINGS.add_source_columns:
                values.extend([source_file, sheet_name, imported_at])
            prepared_rows.append(values)
        database.executemany(sql, prepared_rows)
        log_import(database, source_file, sheet_name, file_hash, table_name, "success", len(rows))
        database.commit()
    except Exception:
        database.rollback()
        raise
    print(f"Imported {len(rows)} rows from {source_file} / {sheet_name} into {table_name}.")
    return len(rows)


def is_inside(directory: Path, candidate: Path) -> bool:
    try:
        candidate.resolve().relative_to(directory.resolve())
        return True
    except ValueError:
        return False


def move_staged_file(file_path: Path, target_directory: Path) -> None:
    if not is_inside(SETTINGS.input_dir, file_path):
        return
    destination = target_directory / file_path.name
    if destination.exists():
        destination = target_directory / f"{file_path.stem}-{int(time.time() * 1000)}{file_path.suffix}"
    shutil.move(str(file_path), str(destination))
    print(f"Moved source file to {destination}")


def import_file(file_name: str | Path, table_override: str | None = None) -> int:
    file_path = Path(file_name).expanduser().resolve()
    if file_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"Unsupported file type: {file_path.suffix}")
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    database: sqlite3.Connection | None = None
    try:
        sheets = read_spreadsheet(file_path)
        file_hash = calculate_file_hash(file_path)
        total_rows = 0
        database = sqlite3.connect(SETTINGS.database_path)
        initialize_database(database)
        for sheet_name, matrix in sheets:
            total_rows += import_sheet(
                database, file_path, file_hash, sheet_name, matrix, len(sheets), table_override
            )
        move_staged_file(file_path, SETTINGS.processed_dir)
        return total_rows
    except Exception:
        try:
            move_staged_file(file_path, SETTINGS.failed_dir)
        except Exception as move_error:
            print(f"Could not move failed file: {move_error}", file=sys.stderr)
        raise
    finally:
        if database is not None:
            database.close()


def ensure_directories() -> None:
    for directory in (SETTINGS.input_dir, SETTINGS.processed_dir, SETTINGS.failed_dir,
                      SETTINGS.database_path.parent):
        directory.mkdir(parents=True, exist_ok=True)


def incoming_files() -> list[Path]:
    return sorted(path for path in SETTINGS.input_dir.iterdir()
                  if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS)


def process_file(file_path: Path, table_override: str | None = None) -> None:
    try:
        import_file(file_path, table_override)
    except Exception as error:
        print(f"Import failed for {file_path}: {error}", file=sys.stderr)


def watch_folder(interval: float) -> None:
    print(f"Watching for spreadsheets in {SETTINGS.input_dir}. Press Ctrl+C to stop.")
    known: set[Path] = set()
    try:
        while True:
            current = set(incoming_files())
            for file_path in sorted(current - known):
                previous_size = -1
                for _ in range(10):
                    size = file_path.stat().st_size
                    if size == previous_size:
                        break
                    previous_size = size
                    time.sleep(0.5)
                process_file(file_path)
            known = set(incoming_files())
            time.sleep(interval)
    except KeyboardInterrupt:
        print("\nWatcher stopped.")


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Import spreadsheets or photographed tables into SQLite.")
    parser.add_argument("--file", help="Import one spreadsheet file")
    parser.add_argument("--table", help="Override the destination table name")
    parser.add_argument("--watch", action="store_true", help="Continue watching the incoming folder")
    return parser.parse_args()


def main() -> int:
    arguments = parse_arguments()
    ensure_directories()
    if arguments.file:
        try:
            import_file(arguments.file, arguments.table)
            return 0
        except Exception as error:
            print(f"Import failed: {error}", file=sys.stderr)
            return 1
    for file_path in incoming_files():
        process_file(file_path, arguments.table)
    if arguments.watch:
        watch_folder(SETTINGS.watch_interval_seconds)
    else:
        print(f"Import complete. Database: {SETTINGS.database_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
