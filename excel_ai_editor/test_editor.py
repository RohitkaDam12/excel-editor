"""
test_editor.py - Automated test suite for the AI Excel Editor prototype.

Tests all 6 specified benchmark prompts:
  1. "Highlight the first row in yellow"
  2. "Make the first row bold"
  3. "Delete the last column"
  4. "Highlight the Email column"
  5. "Make the Email column width 30"
  6. "Highlight the first row and delete the last column"

Verifies both the execution trace output and the openpyxl workbook mutations.
"""

import os
import sys
import json
import openpyxl

# Ensure UTF-8 output encoding on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from excel_operations import inspect_workbook, execute_operation
from llm import understand_request
from create_sample_data import create_sample_workbook


def run_benchmark_tests():
    print("\n" + "#" * 60)
    print("#  RUNNING AI-POWERED EXCEL EDITOR BENCHMARK TEST SUITE   #")
    print("#" * 60)

    base_input = "input.xlsx"
    if not os.path.exists(base_input):
        create_sample_workbook(base_input)

    test_cases = [
        {
            "id": 1,
            "prompt": "Highlight the first row in yellow",
            "verify": lambda ws: (
                ws.cell(row=1, column=1).fill.start_color.rgb in ("FFFF00", "00FFFF00", "FFFFFF00"),
                "Row 1 cells filled with yellow"
            )
        },
        {
            "id": 2,
            "prompt": "Make the first row bold",
            "verify": lambda ws: (
                ws.cell(row=1, column=1).font is not None and ws.cell(row=1, column=1).font.bold is True,
                "Row 1 font is bold"
            )
        },
        {
            "id": 3,
            "prompt": "Delete the last column",
            "verify": lambda ws: (
                ws.max_column == 3 and "Status" not in [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)],
                "Last column (Status) deleted, max_column reduced from 4 to 3"
            )
        },
        {
            "id": 4,
            "prompt": "Highlight the Email column",
            "verify": lambda ws: (
                ws.cell(row=2, column=2).fill.start_color.rgb in ("FFFF00", "00FFFF00", "FFFFFF00"),
                "Email column cells filled with yellow"
            )
        },
        {
            "id": 5,
            "prompt": "Make the Email column width 30",
            "verify": lambda ws: (
                abs(float(ws.column_dimensions["B"].width) - 30.0) < 0.1,
                "Column B (Email) width set to 30.0"
            )
        },
        {
            "id": 6,
            "prompt": "Highlight the first row and delete the last column",
            "verify": lambda ws: (
                ws.cell(row=1, column=1).fill.start_color.rgb in ("FFFF00", "00FFFF00", "FFFFFF00")
                and ws.max_column == 3,
                "Row 1 highlighted in yellow AND last column deleted"
            )
        },
        {
            "id": 7,
            "prompt": "bold the name Alice Johnson",
            "verify": lambda ws: (
                ws.cell(row=2, column=1).font is not None and ws.cell(row=2, column=1).font.bold is True,
                "Cell containing 'Alice Johnson' is bold"
            )
        },
        {
            "id": 8,
            "prompt": "add a new column called summary next to Email",
            "verify": lambda ws: (
                ws.cell(row=1, column=3).value == "summary",
                "Column 'summary' inserted right after 'Email' at column 3"
            )
        },
        {
            "id": 9,
            "prompt": "add a new row at the bottom",
            "verify": lambda ws: (
                ws.max_row == 7,
                "New row appended at the bottom, max_row increased to 7"
            )
        },
        {
            "id": 10,
            "prompt": "delete rows where Status is Inactive",
            "verify": lambda ws: (
                ws.max_row == 5 and all(ws.cell(row=r, column=4).value != "Inactive" for r in range(2, ws.max_row + 1)),
                "Row with Status 'Inactive' deleted, no Inactive rows remain"
            )
        },
        {
            "id": 11,
            "prompt": "delete the Company column",
            "verify": lambda ws: (
                ws.max_column == 3 and all(ws.cell(row=1, column=c).value != "Company" for c in range(1, ws.max_column + 1)),
                "Company column deleted, max_column reduced to 3"
            )
        },
        {
            "id": 12,
            "prompt": "sort rows by Company descending",
            "verify": lambda ws: (
                ws.cell(row=2, column=3).value == "Wayne Enterprises" and ws.cell(row=6, column=3).value == "Acme Corp",
                "Rows sorted by Company descending: top is Wayne Enterprises, bottom is Acme Corp"
            )
        },
        {
            "id": 13,
            "prompt": "arrange the columns alphabetically",
            "verify": lambda ws: (
                [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)] == ["Company", "Email", "Name", "Status"],
                "Columns arranged in alphabetical order: Company, Email, Name, Status"
            )
        },
        {
            "id": 14,
            "prompt": "Convert this sheet into a formatted Excel table",
            "verify": lambda ws: (
                len(ws.tables) >= 1 and "TableStyle" in list(ws.tables.values())[0].tableStyleInfo.name,
                "Excel native Table created with auto-filters and TableStyle"
            )
        },
        {
            "id": 15,
            "prompt": "create a bar chart of Company and Status",
            "verify": lambda ws: (
                len(ws._charts) >= 1,
                "Native Excel Chart embedded onto worksheet"
            )
        },
        {
            "id": 16,
            "prompt": "Copy all rows where Status is Active to a new sheet called Active Records",
            "verify": lambda ws: (
                "Active Records" in getattr(ws, "parent", None).sheetnames
                and getattr(ws, "parent", None)["Active Records"].max_row == 4,
                "Active records copied into new sheet 'Active Records' with headers (4 rows total)"
            )
        },
        {
            "id": 17,
            "prompt": "Export rows where Status is Pending into a new excel file called pending_test.xlsx",
            "verify": lambda ws: (
                os.path.exists("pending_test.xlsx") and openpyxl.load_workbook("pending_test.xlsx").active.max_row == 2,
                "Pending records exported into brand new excel file 'pending_test.xlsx'"
            )
        }
    ]

    all_passed = True

    for test in test_cases:
        t_id = test["id"]
        prompt = test["prompt"]
        out_file = f"test_output_{t_id}.xlsx"

        print(f"\n{'=' * 60}")
        print(f"TEST CASE {t_id}: \"{prompt}\"")
        print(f"{'=' * 60}")

        print("USER REQUEST")
        print(f"Instruction: {prompt}")

        # Load fresh copy of workbook
        wb = openpyxl.load_workbook(base_input)
        ws = wb.active

        # Inspect workbook
        inspection = inspect_workbook(ws)

        # Call LLM
        print("|\nv\nLLM JSON")
        llm_result = understand_request(prompt, headers=inspection["headers"])
        print(json.dumps(llm_result, indent=2))

        # Action Router & Openpyxl Execution
        print("|\nv\nPYTHON ACTION")
        for op in llm_result["operations"]:
            print(f"Action: {op.get('action')}")

        print("|\nv\nOPENPYXL OPERATION")
        for op in llm_result["operations"]:
            success = execute_operation(ws, op, wb=wb)
            if not success:
                print(f"[FAIL] Operation {op} failed execution.")
                all_passed = False

        # Save output file
        print("|\nv\nOUTPUT FILE")
        wb.save(out_file)
        print(f"Saved: {os.path.abspath(out_file)}")

        # Verify assertions
        passed, desc = test["verify"](ws)
        if passed:
            print(f"[PASS] VERIFICATION PASSED: {desc}")
        else:
            print(f"[FAIL] VERIFICATION FAILED: {desc}")
            all_passed = False

    print("\n" + "=" * 60)
    if all_passed:
        print("ALL 17 BENCHMARK TESTS PASSED SUCCESSFULLY! [SUCCESS]")
    else:
        print("SOME TESTS FAILED.")
    print("=" * 60)
    return all_passed


if __name__ == "__main__":
    success = run_benchmark_tests()
    sys.exit(0 if success else 1)
