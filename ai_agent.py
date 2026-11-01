
import os
import sys
import json
import sqlite3
import subprocess
from pathlib import Path

from openai import OpenAI
from import_agent import load_env_file

# Configuration
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "data" / "imports.db"
IMPORT_SCRIPT = BASE_DIR / "import_agent.py"

load_env_file(BASE_DIR / ".env")

client = OpenAI()
MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1")


# Function 1: List SQLite tables
def list_tables():
    if not DB_PATH.exists():
        return {"error": "Database does not exist"}

    with sqlite3.connect(str(DB_PATH)) as conn:
        rows = conn.execute("""
            SELECT name
            FROM sqlite_master
            WHERE type = 'table'
              AND name NOT LIKE 'sqlite_%'
            ORDER BY name
        """).fetchall()

    return {"tables": [row[0] for row in rows]}


# Function 2: Count rows in a table
def count_rows(table_name):
    if not DB_PATH.exists():
        return {"error": "Database does not exist"}

    with sqlite3.connect(str(DB_PATH)) as conn:
        # Verify table exists before using its name in SQL
        row = conn.execute("""
            SELECT name FROM sqlite_master
            WHERE type = 'table' AND name = ?
        """, (table_name,)).fetchone()

        if row is None:
            return {"error": "Table not found: " + table_name}

        # Safely quote the verified SQLite identifier
        safe_name = '"' + row[0].replace('"', '""') + '"'

        count = conn.execute(
            "SELECT COUNT(*) FROM " + safe_name
        ).fetchone()[0]

    return {
        "table": table_name,
        "row_count": count
    }


# Function 3: Run existing importer
def import_files():
    if not IMPORT_SCRIPT.exists():
        return {"error": "import_agent.py not found"}

    try:
        result = subprocess.run(
            [sys.executable, str(IMPORT_SCRIPT)],
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            timeout=300
        )

        return {
            "success": result.returncode == 0,
            "stdout": result.stdout[-5000:],
            "stderr": result.stderr[-2000:]
        }

    except subprocess.TimeoutExpired:
        return {"error": "Import timed out after 300 seconds"}


# Function 4: Display SQLite table contents
def display_table(table_name, limit=100):
    if not DB_PATH.exists():
        return {"error": "Database does not exist"}

    with sqlite3.connect(str(DB_PATH)) as conn:
        # Verify the table exists
        row = conn.execute("""
            SELECT name FROM sqlite_master
            WHERE type = 'table' AND name = ?
        """, (table_name,)).fetchone()

        if row is None:
            return {"error": "Table not found: " + table_name}

        # Safely quote the verified table name
        safe_name = '"' + row[0].replace('"', '""') + '"'

        # Retrieve columns and rows
        cursor = conn.execute(
            f"SELECT * FROM {safe_name} LIMIT ?",
            (max(1, min(int(limit), 1000)),)
        )

        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()

    return {
        "table": table_name,
        "columns": columns,
        "row_count": len(rows),
        "rows": [
            dict(zip(columns, row))
            for row in rows
        ]
    }

# Describe functions to the OpenAI model
TOOLS = [
    {
        "type": "function",
        "name": "list_tables",
        "description": "List all tables in the SQLite database",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": [],
            "additionalProperties": False
        },
        "strict": True
    },
    {
        "type": "function",
        "name": "count_rows",
        "description": "Count rows in a SQLite table",
        "parameters": {
            "type": "object",
            "properties": {
                "table_name": {
                    "type": "string",
                    "description": "Exact SQLite table name"
                }
            },
            "required": ["table_name"],
            "additionalProperties": False
        },
        "strict": True
    },
    {
        "type": "function",
        "name": "import_files",
        "description": (
            "Import supported files from the configured "
            "incoming folder using the existing Python importer"
        ),
        "parameters": {
            "type": "object",
            "properties": {},
            "required": [],
            "additionalProperties": False
        },
        "strict": True
    },
    {
        "type": "function",
        "name": "display_table",
        "description": (
            "Display the actual contents of a SQLite table, "
            "including column names and row values"
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "table_name": {
                    "type": "string",
                    "description": "Exact SQLite table name"
                },
                "limit": {
                    "type": "integer",
                    "description": "Maximum number of rows to display"
                }
            },
            "required": ["table_name", "limit"],
            "additionalProperties": False
        },
        "strict": True
    }
]


FUNCTIONS = {
    "list_tables": list_tables,
    "count_rows": count_rows,
    "import_files": import_files,
    "display_table": display_table
}


def ask_agent(question):
    response = client.responses.create(
        model=MODEL,
        instructions=(
            "You are a spreadsheet import assistant. "
            "Use the provided functions to inspect or import data. "
            "Never invent database results. "
            "Use list_tables to discover table names when needed."
        ),
        input=question,
        tools=TOOLS
    )

    # Allow several rounds of function calls
    for _ in range(8):
        calls = [
            item for item in response.output
            if item.type == "function_call"
        ]

        if not calls:
            return response.output_text

        outputs = []

        for call in calls:
            try:
                args = json.loads(call.arguments)

                if call.name not in FUNCTIONS:
                    result = {"error": "Unknown function"}
                else:
                    result = FUNCTIONS[call.name](**args)

            except Exception as exc:
                result = {"error": str(exc)}

            outputs.append({
                "type": "function_call_output",
                "call_id": call.call_id,
                "output": json.dumps(result)
            })

        response = client.responses.create(
            model=MODEL,
            previous_response_id=response.id,
            input=outputs,
            tools=TOOLS
        )

    return "Stopped: too many function-call rounds."


if __name__ == "__main__":
    print("OpenAI Spreadsheet Agent")
    print("Type 'exit' to quit.")

    while True:
        try:
            question = input("\nYou: ").strip()
        except (EOFError, KeyboardInterrupt):
            break

        if question.lower() in ("exit", "quit"):
            break

        if not question:
            continue

        try:
            answer = ask_agent(question)
            print("\nAgent:", answer)
        except Exception as exc:
            print("\nError:", exc)
