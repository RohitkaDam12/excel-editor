# AI-Powered Excel Editor Prototype (Powered by Google Gemini)

A simple, safe, and beginner-friendly Python prototype of an AI-powered Excel editor.

---

## 💡 Core Concept & Philosophy

1. **The LLM does NOT edit Excel.**
2. **The LLM does NOT generate Python code.**
3. **The LLM ONLY returns structured JSON.**
4. **Python contains all actual Excel manipulation logic.**
5. **The Python Action Router validates and executes predefined openpyxl functions.**
6. **OpenPyXL performs the actual workbook modification.**

```text
User Request
     │
     ▼
Natural-Language Prompt (e.g., "Highlight the first row in yellow")
     │
     ▼
Google Gemini LLM Instruction Parser (llm.py & prompts.py)
     │
     ▼
Structured JSON
{
  "operations": [
    {
      "action": "highlight_row",
      "row": 1,
      "color": "FFFF00"
    }
  ]
}
     │
     ▼
Python Action Router (excel_operations.py)
     │
     ▼
Predefined OpenPyXL Function: highlight_row(ws, 1, "FFFF00")
     │
     ▼
Modified Excel File (output.xlsx)
```

---

## 📁 Project Structure

```text
excel_ai_editor/
│
├── main.py                  # CLI entry point and execution coordinator
├── llm.py                   # Isolated Gemini / LLM client & structured JSON parser
├── excel_operations.py      # Predefined openpyxl functions & Action Router
├── prompts.py               # Strict instruction prompt & JSON schema
├── create_sample_data.py    # Generates initial input.xlsx dataset
├── test_editor.py           # Automated test suite for benchmark instructions
├── requirements.txt         # Dependencies (openpyxl, google-genai, openai, python-dotenv)
├── .env                     # Your local API key configuration
├── .env.example             # Example environment variable file
├── input.xlsx               # Sample workbook (Name, Email, Company, Status)
└── output.xlsx              # Resulting modified workbook
```

---

## 🚀 Setup & API Key Configuration

### 1. Free Google Gemini API Key
You can get a free Gemini API key in seconds from **Google AI Studio**:
👉 [https://aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)

### 2. Configure `.env`
Open the `.env` file in `excel_ai_editor/` and paste your key:
```env
GEMINI_API_KEY=AIzaSy...
GEMINI_MODEL=gemini-2.5-flash
```

*(Note: If no API key is provided, the application automatically uses a built-in offline parser so you can test all sample instructions immediately without an API key or network connection).*

---

## 🖥️ Usage

### Interactive CLI
Run the main program:
```bash
python main.py
```

**Example interaction**:
```text
=============================================
     AI-Powered Excel Editor Prototype       
=============================================
Enter Excel file path [default: input.xlsx]: input.xlsx

What do you want to change?
> Highlight the first row in yellow

==================================================
USER REQUEST
==================================================
Instruction: Highlight the first row in yellow
Input file:  input.xlsx

--- Workbook Inspection ---
Sheet: Sheet1
Rows: 6
Columns: 4

Headers:
1. Name
2. Email
3. Company
4. Status
---------------------------

|
v
LLM JSON
--------------------------------------------------
LLM generated operations:
{
  "operations": [
    {
      "action": "highlight_row",
      "row": 1,
      "color": "FFFF00"
    }
  ]
}

|
v
PYTHON ACTION & OPENPYXL OPERATION
--------------------------------------------------

Router dispatching: highlight_row

Executing operation: highlight_row
  -> Successfully highlighted row 1 with color FFFF00

|
v
OUTPUT FILE
--------------------------------------------------

Saved modified workbook to:
output.xlsx
Successfully applied 1/1 operation(s).
```

### Non-Interactive CLI Arguments
You can also run directly from scripts or command lines:
```bash
python main.py --file input.xlsx --instruction "Make the first row bold" --output output.xlsx
```

---

## 🧪 Automated Benchmark Testing

Run the automated test suite covering all 17 benchmark operations:
```bash
python test_editor.py
```

Benchmark prompts tested:
1. `"Highlight the first row in yellow"`
2. `"Make the first row bold"`
3. `"Delete the last column"`
4. `"Highlight the Email column"`
5. `"Make the Email column width 30"`
6. `"Highlight the first row and delete the last column"`
7. `"bold the name Alice Johnson"`
8. `"add a new column called summary next to Email"`
9. `"add a new row at the bottom"`
10. `"delete rows where Status is Inactive"`
11. `"delete the Company column"`
12. `"sort rows by Company descending"`
13. `"arrange the columns alphabetically"`
14. `"Convert this sheet into a formatted Excel table"`
15. `"create a bar chart of Company and Status"`
16. `"Copy all rows where Status is Active to a new sheet called Active Records"`
17. `"Export rows where Status is Pending into a new excel file called pending_test.xlsx"`

---

## 🛡️ Supported Operations & Validation

| Operation | Parameters | Description |
|---|---|---|
| `highlight_row` | `row`, `color` | Fills cells in row with specified color |
| `highlight_column` | `column`, `color` | Fills cells in column with specified color |
| `highlight_rows_where` | `column`, `value`, `color` | Highlights entire rows where column satisfies condition (e.g. `Age < 25`, `City == Pune`) |
| `highlight_value` | `value`, `column`, `color` | Highlights specific cell(s) matching a value |
| `bold_row` | `row` | Bolds all text in row |
| `bold_column` | `column` | Bolds all text in specified column |
| `bold_cell` | `cell` | Bolds a specific cell (e.g. "B14") |
| `bold_value` | `value`, `column` (opt) | Finds text/name (e.g. "Vikram Joshi") and makes that cell bold |
| `delete_row` | `row` | Deletes a specific row by index or "last" |
| `delete_rows_where` | `column`, `value` / `values` | Deletes rows where column satisfies condition or matches value(s) (bottom-to-top safe) |
| `delete_column` | `column` | Deletes column by name, letter, or "last" |
| `resize_column` | `column`, `width` | Sets width on target column |
| `insert_row` | `row` or `after_row`, `values` | Inserts a new row at or after a position |
| `append_row` | `values` | Appends a new row at the bottom |
| `insert_column` | `after_column` or `column`, `header` | Inserts a new column next to another with a header |
| `update_value_where` | `search_value`, `target_column`, `new_value` | Finds row by search value and updates target column |
| `update_cell` | `cell`, `value` | Directly updates cell value |
| `sort_rows` / `arrange_rows` | `by_column`, `order` | Sorts rows ascending/descending by any parameter (preserves styles) |
| `sort_columns` / `arrange_columns` | `order` | Sorts all columns alphabetically A-Z or Z-A |
| `reorder_columns` | `columns` or `move_column` | Custom column order or moves a column before/after another |
| `create_table` | `name`, `ref`, `style` | Converts data into official native Excel Table with auto-filter dropdowns and banded rows |
| `create_chart` | `chart_type`, `values_column`, `categories_column`, `title`, `position` | Generates embedded native Excel vector charts (bar, line, pie, area) with automatic data aggregation |
| `copy_rows_to_sheet` | `target_sheet`, `column`, `value` / `values` | Copies matching rows (with headers, styles, column widths) to a new or existing sheet |
| `move_rows_to_sheet` | `target_sheet`, `column`, `value` / `values` | Moves matching rows to another sheet (copies to target and deletes from source) |
| `export_to_file` | `output_file`, `column`, `value` / `values` | Filters matching rows and exports them into a brand new standalone Excel workbook |
| `create_sheet` | `sheet_name`, `title` | Creates a new worksheet inside the current workbook |
| `duplicate_sheet` | `source_sheet`, `new_name` | Duplicates an existing sheet |
| `rename_sheet` | `sheet_name`, `new_name` | Renames a sheet in the workbook |
| `delete_sheet` | `sheet_name` | Deletes a sheet from the workbook |

### Smart Resolvers:
- **"last" column / row**: Evaluated dynamically via `ws.max_column` and `ws.max_row`.
- **Header lookup**: Column names specified naturally (e.g. `"Salary"`, `"first_name"`, `"First Name"`) are resolved automatically against Row 1.
- **Auto-Aggregation**: Large datasets (like 500 rows) are aggregated into clean category summaries so charts display clean bars rather than hundreds of unreadable data points.
- **Multi-Sheet & Workbook Exports**: Automatically creates target sheets/files, preserves column headers and cell styling, and handles multiple criteria.
- **Safe File Saving**: If Microsoft Excel holds a lock on `output.xlsx`, the system saves automatically to `output_1.xlsx`, `output_2.xlsx`, etc.
- **Validation**: If an operation references a non-existent column or invalid parameter, validation fails safely and the workbook is **never** corrupted or partially modified.

