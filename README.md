# Spreadsheet Import Agent — Python Version

This Python agent imports Excel (`.xlsx`), CSV, TSV, and photographed tables (`.jpg`, `.jpeg`, `.png`, `.heic`, `.heif`) into SQLite for reporting and further processing. It does not require Node.js or npm.

Regular spreadsheet imports do not use an OpenAI API key. Photo imports use an image-capable OpenAI model by default because photographed screens can defeat traditional OCR. Only photo imports use API tokens.

The earlier npm `xlsx` dependency has been removed. Legacy `.xls` files are not supported; open them in Excel and save them as `.xlsx` first.

## What it does

1. Reads spreadsheet or image files from `incoming`.
2. Uses the first spreadsheet row as column names.
3. Cleans table and column names for SQLite.
4. Infers `INTEGER`, `REAL`, or `TEXT` column types.
5. Creates one SQLite table per worksheet.
6. Inserts all rows inside a transaction.
7. Adds newly discovered columns when a table already exists.
8. Records results in `_spreadsheet_import_log` and skips exact duplicate imports.
9. Moves successful files to `processed` and failed files to `failed`.

Imported rows also contain `_source_file`, `_source_sheet`, and `_imported_at`.

## Requirements

- Python 3.10 or newer
- `pip`

## Install on macOS

```bash
cd spreadsheet-import-agent
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
cp .env.example .env
```

If you are updating an existing installation, reinstall the requirements so the type-annotation compatibility package is added:

```bash
source .venv/bin/activate
python3 -m pip install --upgrade -r requirements.txt
```

Do not share `.venv`. Share `requirements.txt` and let each user install it.

## Import spreadsheets once

Copy `.xlsx`, `.csv`, `.tsv`, `.jpg`, `.jpeg`, `.png`, `.heic`, or `.heif` files into `incoming`, then run:

```bash
python3 import_agent.py
```

For `sample-import-data.xlsx`, the worksheets create:

```text
sample_import_data_pack_data
sample_import_data_employee_rates
```

## Import one file into named tables

```bash
python3 import_agent.py --file "/full/path/sample-import-data.xlsx" --table sample
```

For a multi-sheet workbook, the sheet name is appended, producing tables such as `sample_pack_data` and `sample_employee_rates`.

Files passed with `--file` stay in their original location. Files placed in `incoming` move to `processed` after a successful import.

## Import a file directly from a URL

Download a CSV, TSV, or XLSX file from any public URL and import it in one step:

```bash
python3 import_agent.py --url "https://example.com/data.csv"
```

Combine with `--table` to control the destination table name:

```bash
python3 import_agent.py --url "https://example.com/data.csv" --table my_table
```

The file is downloaded to a temporary folder, imported, and the temporary file is deleted automatically. The `--url` flag cannot be combined with `--file` or `--watch`.

## Watch continuously

```bash
python3 import_agent.py --watch
```

Leave the command running. Press `Ctrl+C` to stop it.

## Import a photographed table

Install the packages and add your API key to `.env`:

```env
IMAGE_OCR_ENGINE=openai
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL=gpt-6-luna
```

Then import the photo:

```bash
python3 import_agent.py --file "/full/path/IMG_2382.jpg" --table prior_day_rates
```

An iPhone HEIC photo works the same way:

```bash
python3 import_agent.py --file "/full/path/table-photo.HEIC" --table prior_day_rates
```

HEIC/HEIF photos are converted to temporary JPEG files before extraction. The original photo is not modified, and the temporary JPEG is deleted automatically.

The image is sent to the configured OpenAI model for table extraction. The resulting rows include a `review_required` column set to `1`; compare the imported values with the original photo before using them for decisions.

### Local OCR option

For photos that must remain entirely on your computer, install Tesseract:

```bash
brew install tesseract
```

Then set:

```env
IMAGE_OCR_ENGINE=tesseract
```

Local OCR uses no API key, but it is less accurate for angled photos, glare, small text, and screen moiré. The agent refuses to load a photo when average OCR confidence is below `OCR_ACCEPTANCE_CONFIDENCE`, preventing obviously unreliable data from silently entering SQLite.

## Inspect the database

The first run creates `data/imports.db`.

```bash
sqlite3 data/imports.db ".tables"
sqlite3 -header -column data/imports.db "SELECT * FROM sample_import_data_pack_data;"
sqlite3 -header -column data/imports.db "SELECT * FROM _spreadsheet_import_log ORDER BY id DESC;"
```

## Configuration

Edit `.env`:

| Setting | Default | Meaning |
|---|---|---|
| `INPUT_DIR` | `incoming` | Folder scanned or watched |
| `PROCESSED_DIR` | `processed` | Successful source-file archive |
| `FAILED_DIR` | `failed` | Failed source-file archive |
| `DATABASE_PATH` | `data/imports.db` | SQLite database file |
| `ADD_SOURCE_COLUMNS` | `true` | Add source metadata to imported rows |
| `WATCH_INTERVAL_SECONDS` | `2` | Folder polling interval |
| `IMAGE_OCR_ENGINE` | `openai` | `openai` for difficult photos or `tesseract` for local OCR |
| `OPENAI_API_KEY` | blank | Required only when importing photos with OpenAI |
| `OPENAI_MODEL` | `gpt-6-luna` | Image-capable model used for photo extraction |
| `TESSERACT_COMMAND` | `tesseract` | Local OCR executable |
| `OCR_LANGUAGE` | `eng` | Local OCR language |
| `OCR_MIN_CONFIDENCE` | `20` | Ignore very weak OCR words |
| `OCR_ACCEPTANCE_CONFIDENCE` | `60` | Reject low-confidence local OCR imports |

Relative paths are resolved from the project folder. Absolute paths also work.

## Notes

- The first row must contain headings.
- Blank headings become `column_1`, `column_2`, and so on.
- Duplicate headings receive suffixes such as `_2`.
- `.xlsx` formulas use their last calculated values saved by Excel.
- CSV and TSV values are imported as text because those formats do not preserve data types.
- The file hash, worksheet, and table name prevent an exact duplicate import.
- Always review data extracted from a photo. Neither traditional OCR nor AI vision is guaranteed to transcribe every cell correctly.


