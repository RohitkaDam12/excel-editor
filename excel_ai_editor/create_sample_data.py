"""
create_sample_data.py - Helper script to generate input.xlsx sample data.

Creates an Excel spreadsheet with columns:
Name | Email | Company | Status
and 5 rows of realistic sample records.
"""

import os
import openpyxl


def create_sample_workbook(file_path: str = "input.xlsx") -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Header row
    headers = ["Name", "Email", "Company", "Status"]
    ws.append(headers)

    # 5 rows of sample data
    rows = [
        ["Alice Johnson", "alice@example.com", "Acme Corp", "Active"],
        ["Bob Smith", "bob@example.com", "Globex Inc", "Pending"],
        ["Charlie Brown", "charlie@example.com", "Initech LLC", "Active"],
        ["Diana Prince", "diana@example.com", "Wayne Enterprises", "Inactive"],
        ["Evan Wright", "evan@example.com", "Stark Industries", "Active"],
    ]

    for row in rows:
        ws.append(row)

    wb.save(file_path)
    print(f"Sample workbook successfully created at: {os.path.abspath(file_path)}")
    return file_path


if __name__ == "__main__":
    create_sample_workbook("input.xlsx")
