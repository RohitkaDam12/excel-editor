"""
excel_operations.py - Predefined Excel operations and Action Router.

This file contains ALL actual Excel manipulation logic using openpyxl.
The LLM never generates or executes Python code directly;
instead, the router safely validates and dispatches JSON operations
to the trusted functions defined below.
"""

import datetime
from copy import copy
import re
from typing import Any
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, LineChart, PieChart, AreaChart, Reference

# Set of operations supported by this prototype
ALLOWED_ACTIONS = {
    "highlight_row",
    "highlight_column",
    "highlight_rows_where",
    "highlight_value",
    "bold_row",
    "bold_column",
    "bold_cell",
    "bold_value",
    "delete_row",
    "delete_rows_where",
    "delete_column",
    "resize_column",
    "insert_row",
    "append_row",
    "insert_column",
    "update_value_where",
    "update_cell",
    "sort_rows",
    "arrange_rows",
    "sort_columns",
    "arrange_columns",
    "reorder_columns",
    "create_table",
    "create_chart",
    "copy_rows_to_sheet",
    "move_rows_to_sheet",
    "create_sheet",
    "duplicate_sheet",
    "rename_sheet",
    "delete_sheet",
    "export_to_file",
}

# Mapping of common color names to standard aRGB hex codes
COLOR_MAP = {
    "yellow": "FFFF00",
    "light_yellow": "FFFFE0",
    "red": "FF0000",
    "light_red": "FFC7CE",
    "green": "00FF00",
    "light_green": "C6EFCE",
    "blue": "0000FF",
    "light_blue": "BDD7EE",
    "orange": "FFA500",
    "purple": "800080",
    "gray": "808080",
    "grey": "808080",
    "cyan": "00FFFF",
    "magenta": "FF00FF",
    "white": "FFFFFF",
    "black": "000000",
}


def standardize_color(color_input: Any) -> str:
    """
    Standardize color input (name or hex) to a valid 6-char or 8-char hex string.
    Defaults to yellow ('FFFF00') if invalid or missing.
    """
    if not color_input:
        return "FFFF00"

    cleaned = str(color_input).strip().lower().replace(" ", "_")
    if cleaned in COLOR_MAP:
        return COLOR_MAP[cleaned]

    # Handle hex codes (with or without '#')
    hex_val = str(color_input).strip().lstrip("#").upper()
    if len(hex_val) == 6 or len(hex_val) == 8:
        return hex_val

    # Default fallback
    return "FFFF00"


def inspect_workbook(ws: openpyxl.worksheet.worksheet.Worksheet) -> dict[str, Any]:
    """
    Inspects and prints sheet metadata: title, row count, column count, and header row.
    """
    sheet_name = ws.title
    row_count = ws.max_row or 0
    col_count = ws.max_column or 0

    headers: list[str] = []
    if row_count >= 1:
        for col in range(1, col_count + 1):
            val = ws.cell(row=1, column=col).value
            headers.append(str(val) if val is not None else f"Column {col}")

    print("\n--- Workbook Inspection ---")
    print(f"Sheet: {sheet_name}")
    print(f"Rows: {row_count}")
    print(f"Columns: {col_count}")
    print("\nHeaders:")
    for idx, header in enumerate(headers, start=1):
        print(f"{idx}. {header}")
    print("---------------------------\n")

    return {
        "sheet_name": sheet_name,
        "rows": row_count,
        "columns": col_count,
        "headers": headers,
    }


def resolve_column_index(ws: openpyxl.worksheet.worksheet.Worksheet, column: Any) -> int:
    """
    Resolves a column reference to a 1-based column integer index.
    Supported references:
      - "last" -> ws.max_column
      - integer or string integer (e.g. 2, "2")
      - Excel column letter (e.g. "B", "AA")
      - Header name (e.g. "Email", "Name") matched against row 1
    """
    if ws.max_column is None or ws.max_column < 1:
        raise ValueError("Worksheet has no columns.")

    # 1. Handle "last"
    if str(column).strip().lower() == "last":
        return ws.max_column

    # 2. Handle direct integer or numeric string
    if isinstance(column, int):
        if 1 <= column <= ws.max_column:
            return column
        raise ValueError(f"Column index {column} is out of bounds (1 to {ws.max_column}).")

    col_str = str(column).strip()
    if col_str.isdigit():
        val = int(col_str)
        if 1 <= val <= ws.max_column:
            return val
        raise ValueError(f"Column index {val} is out of bounds (1 to {ws.max_column}).")

    # 3. Handle Header Name lookup in Row 1 (first row)
    # First pass: exact (case-insensitive) match
    for col_idx in range(1, ws.max_column + 1):
        header_val = ws.cell(row=1, column=col_idx).value
        if header_val is not None:
            val_clean = str(header_val).strip().lower()
            if val_clean == col_str.lower():
                return col_idx

    # Second pass: normalized underscore/space match (e.g., 'raw_notes' vs 'raw notes')
    for col_idx in range(1, ws.max_column + 1):
        header_val = ws.cell(row=1, column=col_idx).value
        if header_val is not None:
            norm_val = str(header_val).strip().lower().replace("_", " ")
            norm_col = col_str.lower().replace("_", " ")
            if norm_val == norm_col:
                return col_idx

    # 4. Handle Excel column letter (e.g., "A", "B", "C", "AA")
    if col_str.isalpha():
        try:
            letter_idx = column_index_from_string(col_str)
            if 1 <= letter_idx <= ws.max_column:
                return letter_idx
        except ValueError:
            pass

    # Third pass: partial / suffix match if distinct (e.g., 'notes' matching 'raw_notes')
    matches = []
    for col_idx in range(1, ws.max_column + 1):
        header_val = ws.cell(row=1, column=col_idx).value
        if header_val is not None:
            norm_val = str(header_val).strip().lower()
            if col_str.lower() in norm_val.split("_") or col_str.lower() in norm_val.split(" "):
                matches.append(col_idx)
    if len(matches) == 1:
        return matches[0]

    raise ValueError(f"Column '{column}' was not found in the workbook.")


def resolve_row_number(ws: openpyxl.worksheet.worksheet.Worksheet, row: Any) -> int:
    """
    Resolves a row reference to a 1-based row integer.
    Supports integers, string integers, "first" (1), and "last" (ws.max_row).
    """
    if ws.max_row is None or ws.max_row < 1:
        raise ValueError("Worksheet has no rows.")

    row_str = str(row).strip().lower()
    if row_str == "first":
        return 1
    if row_str == "last":
        return ws.max_row

    try:
        row_num = int(row)
    except (ValueError, TypeError):
        raise ValueError(f"Invalid row number: '{row}'.")

    if row_num < 1 or row_num > ws.max_row:
        raise ValueError(f"Row {row_num} is out of bounds (1 to {ws.max_row}).")

    return row_num


# =====================================================================
# Predefined OpenPyXL Functions
# =====================================================================

def highlight_row(ws: openpyxl.worksheet.worksheet.Worksheet, row_number: int, color: str = "FFFF00") -> None:
    """
    Highlights every cell in the specified row using a solid PatternFill.
    """
    resolved_row = resolve_row_number(ws, row_number)
    hex_color = standardize_color(color)
    fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

    max_col = ws.max_column or 1
    for col_idx in range(1, max_col + 1):
        ws.cell(row=resolved_row, column=col_idx).fill = fill


def highlight_column(ws: openpyxl.worksheet.worksheet.Worksheet, column: Any, color: str = "FFFF00") -> None:
    """
    Highlights every cell in the specified column using a solid PatternFill.
    """
    col_idx = resolve_column_index(ws, column)
    hex_color = standardize_color(color)
    fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

    max_row = ws.max_row or 1
    for row_idx in range(1, max_row + 1):
        ws.cell(row=row_idx, column=col_idx).fill = fill


def parse_operator_and_value(value: Any, operator: str | None = None) -> tuple[str, Any]:
    """
    Extracts operator and target value.
    Supports symbols (<, <=, >, >=, !=, ==, =) and phrases (less than, greater than, under, etc.).
    """
    if operator:
        op = str(operator).strip()
        if op == "=":
            op = "=="
        return op, value

    if isinstance(value, str):
        v = value.strip()
        for sym in ["<=", ">=", "!=", "<>", "==", "<", ">", "="]:
            if v.startswith(sym):
                op = sym if sym != "<>" else "!="
                if op == "=":
                    op = "=="
                clean_val = v[len(sym):].strip()
                return op, clean_val

        lower_v = v.lower()
        if lower_v.startswith("less than or equal to") or lower_v.startswith("at most"):
            m = re.search(r"(?:less than or equal to|at most)\s*(.+)", lower_v)
            if m:
                return "<=", m.group(1).strip()
        elif lower_v.startswith("greater than or equal to") or lower_v.startswith("at least"):
            m = re.search(r"(?:greater than or equal to|at least)\s*(.+)", lower_v)
            if m:
                return ">=", m.group(1).strip()
        elif lower_v.startswith("less than") or lower_v.startswith("under") or lower_v.startswith("below"):
            m = re.search(r"(?:less than|under|below)\s*(.+)", lower_v)
            if m:
                return "<", m.group(1).strip()
        elif lower_v.startswith("greater than") or lower_v.startswith("over") or lower_v.startswith("above"):
            m = re.search(r"(?:greater than|over|above)\s*(.+)", lower_v)
            if m:
                return ">", m.group(1).strip()
        elif lower_v.startswith("not equal to") or lower_v.startswith("not"):
            m = re.search(r"(?:not equal to|not)\s*(.+)", lower_v)
            if m:
                return "!=", m.group(1).strip()

        elif lower_v.startswith("contains") or lower_v.startswith("containing"):
            m = re.search(r"(?:contains|containing)\s*(.+)", lower_v)
            if m:
                return "contains", m.group(1).strip()
        elif lower_v.startswith("starts with") or lower_v.startswith("starting with"):
            m = re.search(r"(?:starts with|starting with)\s*(.+)", lower_v)
            if m:
                return "starts_with", m.group(1).strip()
        elif lower_v.startswith("ends with") or lower_v.startswith("ending with"):
            m = re.search(r"(?:ends with|ending with)\s*(.+)", lower_v)
            if m:
                return "ends_with", m.group(1).strip()

    return "==", value


def matches_condition(cell_val: Any, op: str, target: Any) -> bool:
    """
    Evaluates whether cell_val satisfies (op target).
    Handles numeric, datetime, and string comparisons.
    """
    if cell_val is None:
        return False

    # 1. Try numeric comparison
    try:
        num_cell = float(cell_val)
        num_target = float(str(target).strip())
        if op == "<":
            return num_cell < num_target
        elif op == "<=":
            return num_cell <= num_target
        elif op == ">":
            return num_cell > num_target
        elif op == ">=":
            return num_cell >= num_target
        elif op in ("==", "="):
            return num_cell == num_target
        elif op in ("!=", "<>"):
            return num_cell != num_target
    except (ValueError, TypeError):
        pass

    # 2. Try date/datetime comparison
    if hasattr(cell_val, "strftime"):
        try:
            if isinstance(target, str):
                target_dt = datetime.datetime.fromisoformat(target.strip())
                if isinstance(cell_val, datetime.date) and not isinstance(cell_val, datetime.datetime):
                    target_dt = target_dt.date()
                if op == "<":
                    return cell_val < target_dt
                elif op == "<=":
                    return cell_val <= target_dt
                elif op == ">":
                    return cell_val > target_dt
                elif op == ">=":
                    return cell_val >= target_dt
                elif op in ("==", "="):
                    return cell_val == target_dt
                elif op in ("!=", "<>"):
                    return cell_val != target_dt
        except Exception:
            pass

    # 3. String comparison
    c_str = str(cell_val).strip().lower()
    t_str = str(target).strip().lower()

    if op == "<":
        return c_str < t_str
    elif op == "<=":
        return c_str <= t_str
    elif op == ">":
        return c_str > t_str
    elif op == ">=":
        return c_str >= t_str
    elif op in ("!=", "<>"):
        return c_str != t_str
    elif op == "contains":
        return t_str in c_str
    elif op == "starts_with":
        return c_str.startswith(t_str)
    elif op == "ends_with":
        return c_str.endswith(t_str)
    else:  # '==' or '='
        return t_str == c_str


def highlight_rows_where(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    column: Any,
    value: Any,
    color: str = "FFFF00",
    operator: str | None = None,
) -> list[int]:
    """
    Highlights entire rows where the specified column satisfies the condition (e.g. value '< 25' or operator='<' value=25).
    Skips header row (row 1). Returns list of highlighted row numbers.
    """
    col_idx = resolve_column_index(ws, column)
    hex_color = standardize_color(color)
    fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
    op, target = parse_operator_and_value(value, operator)

    highlighted_rows: list[int] = []
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1

    for row_idx in range(2, max_row + 1):
        cell_val = ws.cell(row=row_idx, column=col_idx).value
        if cell_val is not None and matches_condition(cell_val, op, target):
            for c in range(1, max_col + 1):
                ws.cell(row=row_idx, column=c).fill = fill
            highlighted_rows.append(row_idx)

    if not highlighted_rows:
        raise ValueError(f"No rows found where '{column}' satisfies '{op} {target}'.")
    return highlighted_rows


def highlight_value(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    value: Any,
    column: Any = None,
    color: str = "FFFF00",
) -> list[str]:
    """
    Highlights individual cell(s) that contain/match value.
    Returns list of highlighted cell coordinates.
    """
    hex_color = standardize_color(color)
    fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
    val_str = str(value).strip().lower()

    if column is not None:
        col_indices = [resolve_column_index(ws, column)]
    else:
        col_indices = list(range(1, (ws.max_column or 1) + 1))

    highlighted_cells: list[str] = []
    max_row = ws.max_row or 1

    for row_idx in range(1, max_row + 1):
        for col_idx in col_indices:
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                c_text = str(cell.value).strip().lower()
                if val_str == c_text or val_str in c_text:
                    cell.fill = fill
                    highlighted_cells.append(cell.coordinate)

    if not highlighted_cells:
        raise ValueError(f"No cell found matching '{value}'.")
    return highlighted_cells


def update_value_where(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    search_value: Any,
    target_column: Any,
    new_value: Any,
    search_column: Any = None,
) -> list[str]:
    """
    Finds rows where search_column (or any column if None) matches search_value,
    and updates target_column with new_value.
    Returns list of updated cell coordinates.
    """
    target_col_idx = resolve_column_index(ws, target_column)
    val_str = str(search_value).strip().lower()

    if search_column is not None:
        search_cols = [resolve_column_index(ws, search_column)]
    else:
        search_cols = list(range(1, (ws.max_column or 1) + 1))

    updated_coords: list[str] = []
    max_row = ws.max_row or 1

    for row_idx in range(2, max_row + 1):
        match_found = False
        for c_idx in search_cols:
            c_val = ws.cell(row=row_idx, column=c_idx).value
            if c_val is not None:
                c_text = str(c_val).strip().lower()
                if val_str == c_text or val_str in c_text:
                    match_found = True
                    break

        if match_found:
            cell = ws.cell(row=row_idx, column=target_col_idx)
            cell.value = new_value
            updated_coords.append(cell.coordinate)

    if not updated_coords:
        raise ValueError(f"No rows found matching '{search_value}' to update '{target_column}'.")
    return updated_coords


def update_cell(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    cell: str,
    value: Any,
) -> str:
    """Updates the value of a specific cell coordinate (e.g. 'J11')."""
    target_coord = str(cell).strip().upper()
    ws[target_coord].value = value
    return target_coord


def bold_row(ws: openpyxl.worksheet.worksheet.Worksheet, row_number: int) -> None:
    """
    Makes every cell in the requested row bold, preserving font family/size if present.
    """
    resolved_row = resolve_row_number(ws, row_number)
    max_col = ws.max_column or 1

    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=resolved_row, column=col_idx)
        if cell.font:
            cell.font = Font(
                name=cell.font.name,
                size=cell.font.size,
                bold=True,
                italic=cell.font.italic,
                color=cell.font.color,
            )
        else:
            cell.font = Font(bold=True)


def bold_column(ws: openpyxl.worksheet.worksheet.Worksheet, column: Any) -> None:
    """
    Makes every cell in the requested column bold.
    """
    col_idx = resolve_column_index(ws, column)
    max_row = ws.max_row or 1
    for row_idx in range(1, max_row + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        if cell.font:
            cell.font = Font(
                name=cell.font.name,
                size=cell.font.size,
                bold=True,
                italic=cell.font.italic,
                color=cell.font.color,
            )
        else:
            cell.font = Font(bold=True)


def bold_cell(ws: openpyxl.worksheet.worksheet.Worksheet, cell_coord: str) -> None:
    """
    Makes a specific cell coordinate (e.g. 'B14') bold.
    """
    cell = ws[str(cell_coord).strip().upper()]
    if cell.font:
        cell.font = Font(
            name=cell.font.name,
            size=cell.font.size,
            bold=True,
            italic=cell.font.italic,
            color=cell.font.color,
        )
    else:
        cell.font = Font(bold=True)


def bold_value(ws: openpyxl.worksheet.worksheet.Worksheet, value: Any, column: Any = None) -> list[str]:
    """
    Searches for text/value in the worksheet (or within a specific column) and makes matching cell(s) bold.
    Returns the list of coordinates bolded.
    """
    val_str = str(value).strip().lower()
    bolded_cells: list[str] = []

    if column is not None:
        col_indices = [resolve_column_index(ws, column)]
    else:
        col_indices = list(range(1, (ws.max_column or 1) + 1))

    max_row = ws.max_row or 1
    for row_idx in range(1, max_row + 1):
        for col_idx in col_indices:
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                cell_text = str(cell.value).strip().lower()
                if val_str in cell_text or val_str == cell_text:
                    if cell.font:
                        cell.font = Font(
                            name=cell.font.name,
                            size=cell.font.size,
                            bold=True,
                            italic=cell.font.italic,
                            color=cell.font.color,
                        )
                    else:
                        cell.font = Font(bold=True)
                    bolded_cells.append(cell.coordinate)

    if not bolded_cells:
        raise ValueError(f"No cell found matching '{value}'.")
    return bolded_cells


def insert_row(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    row_number: Any = None,
    after_row: Any = None,
    values: list[Any] | None = None,
) -> int:
    """
    Inserts a new row at the given row number or after a given row.
    """
    if after_row is not None:
        target_row = resolve_row_number(ws, after_row) + 1
    elif row_number is not None:
        target_row = resolve_row_number(ws, row_number)
    else:
        target_row = (ws.max_row or 0) + 1

    ws.insert_rows(target_row)

    if values and isinstance(values, list):
        for col_idx, val in enumerate(values, start=1):
            ws.cell(row=target_row, column=col_idx, value=val)

    return target_row


def append_row(ws: openpyxl.worksheet.worksheet.Worksheet, values: list[Any] | None = None) -> int:
    """
    Appends a new row at the end of the sheet.
    """
    row_values = values if (values and isinstance(values, list) and len(values) > 0) else [None]
    ws.append(row_values)
    return ws.max_row


def insert_column(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    column: Any = None,
    after_column: Any = None,
    header: str | None = None,
    values: list[Any] | None = None,
) -> int:
    """
    Inserts a new column at a given column position or after an existing column.
    """
    if after_column is not None:
        base_col = resolve_column_index(ws, after_column)
        target_col = base_col + 1
    elif column is not None:
        col_str = str(column).strip().lower()
        if col_str in ("last", "end"):
            target_col = (ws.max_column or 0) + 1
        else:
            target_col = resolve_column_index(ws, column)
    else:
        target_col = (ws.max_column or 0) + 1

    ws.insert_cols(target_col)

    if header:
        ws.cell(row=1, column=target_col, value=str(header))

    if values and isinstance(values, list):
        for row_idx, val in enumerate(values, start=2):
            ws.cell(row=row_idx, column=target_col, value=val)

    return target_col


def delete_row(ws: openpyxl.worksheet.worksheet.Worksheet, row_number: int) -> None:
    """
    Deletes the specified row using openpyxl.
    """
    resolved_row = resolve_row_number(ws, row_number)
    ws.delete_rows(resolved_row)


def delete_rows_where(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    column: Any = None,
    value: Any = None,
    values: list[Any] | None = None,
    operator: str | None = None,
) -> list[int]:
    """
    Deletes entire rows where the specified column satisfies the condition or matches value(s).
    If column is None, checks across all columns in that row.
    Iterates in reverse (from bottom to top) to ensure row indices do not shift incorrectly during deletion.
    Returns list of original 1-based row numbers that were deleted.
    """
    col_idx = None
    if column is not None and str(column).strip() != "":
        col_idx = resolve_column_index(ws, column)

    # Build list of target values
    if values is not None and isinstance(values, list):
        target_list = values
    elif isinstance(value, list):
        target_list = value
    elif isinstance(value, str) and (" and " in value.lower() or "," in value) and not any(sym in value for sym in ["<", ">", "=", "!"]):
        if " and " in value.lower():
            target_list = [v.strip() for v in re.split(r"\s+and\s+", value, flags=re.IGNORECASE) if v.strip()]
        else:
            target_list = [v.strip() for v in value.split(",") if v.strip()]
    else:
        target_list = [value] if value is not None else []

    parsed_conditions = [parse_operator_and_value(t, operator) for t in target_list]

    deleted_rows: list[int] = []
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1

    # Iterate in reverse from bottom to top, skipping header row 1
    for row_idx in range(max_row, 1, -1):
        matched = False
        if col_idx is not None:
            cell_val = ws.cell(row=row_idx, column=col_idx).value
            if cell_val is not None:
                for op, target in parsed_conditions:
                    if matches_condition(cell_val, op, target):
                        matched = True
                        break
        else:
            for c in range(1, max_col + 1):
                cell_val = ws.cell(row=row_idx, column=c).value
                if cell_val is not None:
                    for op, target in parsed_conditions:
                        if matches_condition(cell_val, op, target):
                            matched = True
                            break
                if matched:
                    break

        if matched:
            ws.delete_rows(row_idx)
            deleted_rows.append(row_idx)

    if not deleted_rows:
        col_desc = f"in column '{column}'" if column is not None else "in any column"
        raise ValueError(f"No rows found satisfying '{value or values}' {col_desc}.")

    return deleted_rows


def delete_column(ws: openpyxl.worksheet.worksheet.Worksheet, column: Any) -> list[int]:
    """
    Deletes the specified column(s) using openpyxl.
    Supports single column, list of columns, or multiple columns.
    Deletes from highest index to lowest to avoid index shifting.
    """
    if isinstance(column, list):
        col_list = column
    elif isinstance(column, str) and (" and " in column.lower() or "," in column):
        if " and " in column.lower():
            col_list = [c.strip() for c in re.split(r"\s+and\s+", column, flags=re.IGNORECASE) if c.strip()]
        else:
            col_list = [c.strip() for c in column.split(",") if c.strip()]
    else:
        col_list = [column]

    indices = []
    for c in col_list:
        idx = resolve_column_index(ws, c)
        if idx not in indices:
            indices.append(idx)

    # Sort descending so deleting does not alter remaining targets
    indices.sort(reverse=True)
    for idx in indices:
        ws.delete_cols(idx)
    return indices


def resize_column(ws: openpyxl.worksheet.worksheet.Worksheet, column: Any, width: Any) -> None:
    """
    Resizes the specified column. Column can be letter, number, 'last', or header name.
    """
    col_idx = resolve_column_index(ws, column)
    col_letter = get_column_letter(col_idx)

    try:
        numeric_width = float(width)
        if numeric_width <= 0:
            raise ValueError()
    except (ValueError, TypeError):
        raise ValueError(f"Invalid column width: '{width}'. Width must be a positive number.")

    ws.column_dimensions[col_letter].width = numeric_width


def sort_rows(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    by_column: Any,
    order: str = "ascending",
    has_header: bool = True,
) -> int:
    """
    Sorts rows (skipping header row 1 if has_header=True) based on values in by_column.
    Supports ascending (asc, a-z, low to high) or descending (desc, z-a, high to low).
    Preserves all cell values and styling (font, fill, border, alignment, number_format).
    Returns the count of sorted rows.
    """
    col_idx = resolve_column_index(ws, by_column)
    start_row = 2 if has_header else 1
    max_row = ws.max_row or 0
    max_col = ws.max_column or 0

    if max_row <= start_row:
        return 0

    order_str = str(order).strip().lower()
    reverse = any(k in order_str for k in ["desc", "reverse", "z to a", "z-a", "high to low", "largest"])

    # Extract row records and their styles
    rows_data = []
    for r in range(start_row, max_row + 1):
        row_cells = []
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            row_cells.append({
                "value": cell.value,
                "font": copy(cell.font) if cell.font else None,
                "fill": copy(cell.fill) if cell.fill else None,
                "border": copy(cell.border) if cell.border else None,
                "alignment": copy(cell.alignment) if cell.alignment else None,
                "number_format": cell.number_format,
            })
        sort_val = row_cells[col_idx - 1]["value"]
        rows_data.append((sort_val, row_cells))

    # Type-safe 3-tuple sort key: (type_group, num_val, str_val)
    # Prevents Python 3 TypeError when comparing different types
    def make_sort_key(reverse_flag):
        def _key(item):
            val = item[0]
            if val is None or str(val).strip() == "":
                return (2 if not reverse_flag else -1, 0.0, "")
            if isinstance(val, (int, float)):
                return (0, float(val), "")
            if isinstance(val, str):
                s = val.strip().replace(",", "")
                try:
                    return (0, float(s), "")
                except ValueError:
                    pass
            if hasattr(val, "isoformat"):
                return (0, 0.0, str(val.isoformat()))
            return (1, 0.0, str(val).strip().lower())
        return _key

    rows_data.sort(key=make_sort_key(reverse), reverse=reverse)

    # Write sorted rows back into worksheet
    for row_offset, (_, row_cells) in enumerate(rows_data):
        r = start_row + row_offset
        for c_idx, cell_data in enumerate(row_cells, start=1):
            cell = ws.cell(row=r, column=c_idx)
            cell.value = cell_data["value"]
            if cell_data["font"]:
                cell.font = cell_data["font"]
            if cell_data["fill"]:
                cell.fill = cell_data["fill"]
            if cell_data["border"]:
                cell.border = cell_data["border"]
            if cell_data["alignment"]:
                cell.alignment = cell_data["alignment"]
            if cell_data["number_format"]:
                cell.number_format = cell_data["number_format"]

    return len(rows_data)


def reorder_columns(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    order: list[Any] | None = None,
    move_column: Any = None,
    before_column: Any = None,
    after_column: Any = None,
    position: Any = None,
) -> list[str]:
    """
    Reorders columns in the worksheet according to a list of columns or by moving a specific column.
    Preserves all cell values, styles, and dimensions.
    Returns list of reordered column header names.
    """
    max_col = ws.max_column or 0
    max_row = ws.max_row or 0
    if max_col <= 1:
        return [str(ws.cell(row=1, column=1).value or "")]

    # Collect current columns data
    current_columns = []
    for c in range(1, max_col + 1):
        col_cells = []
        for r in range(1, max_row + 1):
            cell = ws.cell(row=r, column=c)
            col_cells.append({
                "value": cell.value,
                "font": copy(cell.font) if cell.font else None,
                "fill": copy(cell.fill) if cell.fill else None,
                "border": copy(cell.border) if cell.border else None,
                "alignment": copy(cell.alignment) if cell.alignment else None,
                "number_format": cell.number_format,
            })
        header_val = ws.cell(row=1, column=c).value
        header_str = str(header_val) if header_val is not None else f"Column {c}"
        current_columns.append((c, header_str, col_cells))

    # Determine new ordering of 0-based indices
    new_order_indices = []

    if order and isinstance(order, list) and len(order) > 0:
        for item in order:
            try:
                col_idx = resolve_column_index(ws, item) - 1
                if col_idx not in new_order_indices and 0 <= col_idx < max_col:
                    new_order_indices.append(col_idx)
            except Exception:
                pass
        # Append remaining columns not explicitly in order
        for idx in range(max_col):
            if idx not in new_order_indices:
                new_order_indices.append(idx)

    elif move_column is not None:
        target_idx = resolve_column_index(ws, move_column) - 1
        remaining_indices = [i for i in range(max_col) if i != target_idx]

        if before_column is not None:
            before_idx = resolve_column_index(ws, before_column) - 1
            insert_pos = remaining_indices.index(before_idx) if before_idx in remaining_indices else 0
            remaining_indices.insert(insert_pos, target_idx)
            new_order_indices = remaining_indices
        elif after_column is not None:
            after_idx = resolve_column_index(ws, after_column) - 1
            insert_pos = (remaining_indices.index(after_idx) + 1) if after_idx in remaining_indices else len(remaining_indices)
            remaining_indices.insert(insert_pos, target_idx)
            new_order_indices = remaining_indices
        elif position is not None:
            pos_str = str(position).strip().lower()
            if pos_str in ("first", "start", "1"):
                new_order_indices = [target_idx] + remaining_indices
            elif pos_str in ("last", "end"):
                new_order_indices = remaining_indices + [target_idx]
            else:
                try:
                    num_pos = max(0, min(int(pos_str) - 1, len(remaining_indices)))
                    remaining_indices.insert(num_pos, target_idx)
                    new_order_indices = remaining_indices
                except ValueError:
                    new_order_indices = list(range(max_col))
        else:
            new_order_indices = list(range(max_col))
    else:
        new_order_indices = list(range(max_col))

    # Write reordered columns back to worksheet
    for new_c_idx, orig_col_pos in enumerate(new_order_indices, start=1):
        _, _, col_cells = current_columns[orig_col_pos]
        for r_idx, cell_data in enumerate(col_cells, start=1):
            cell = ws.cell(row=r_idx, column=new_c_idx)
            cell.value = cell_data["value"]
            if cell_data["font"]:
                cell.font = cell_data["font"]
            if cell_data["fill"]:
                cell.fill = cell_data["fill"]
            if cell_data["border"]:
                cell.border = cell_data["border"]
            if cell_data["alignment"]:
                cell.alignment = cell_data["alignment"]
            if cell_data["number_format"]:
                cell.number_format = cell_data["number_format"]

    return [current_columns[i][1] for i in new_order_indices]


def sort_columns(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    order: str = "ascending",
) -> list[str]:
    """
    Sorts all columns alphabetically by their header in row 1.
    """
    max_col = ws.max_column or 0
    if max_col <= 1:
        return [str(ws.cell(row=1, column=1).value or "")]

    headers = []
    for c in range(1, max_col + 1):
        val = ws.cell(row=1, column=c).value
        headers.append(str(val) if val is not None else f"Column {c}")

    order_str = str(order).strip().lower()
    reverse = any(k in order_str for k in ["desc", "reverse", "z to a", "z-a"])
    sorted_headers = sorted(headers, key=lambda h: h.lower(), reverse=reverse)

    return reorder_columns(ws, order=sorted_headers)


def create_table(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    name: str | None = None,
    ref: str | None = None,
    style: str = "TableStyleMedium9",
    show_row_stripes: bool = True,
    show_column_stripes: bool = False,
) -> str:
    """
    Creates an official native Excel Table (ListObject) on the worksheet.
    Adds auto-filters, banded rows, and official Excel table styling.
    Returns the table display name.
    """
    max_col = ws.max_column or 1
    max_row = ws.max_row or 1

    if max_row < 1 or max_col < 1:
        raise ValueError("Worksheet is empty; cannot create table.")

    if not ref:
        end_col_letter = get_column_letter(max_col)
        ref = f"A1:{end_col_letter}{max_row}"

    # Generate or sanitize a valid Excel Table name (alphanumeric/underscore, unique)
    existing_tables = set(ws.tables.keys()) if hasattr(ws, "tables") else set()
    if not name:
        counter = len(existing_tables) + 1
        table_name = f"DataTable_{counter}"
        while table_name in existing_tables:
            counter += 1
            table_name = f"DataTable_{counter}"
    else:
        clean = re.sub(r"[^a-zA-Z0-9_]", "_", str(name).strip())
        if not clean or not (clean[0].isalpha() or clean[0] == "_"):
            clean = f"Table_{clean}"
        base_clean = clean
        counter = 1
        table_name = clean
        while table_name in existing_tables:
            table_name = f"{base_clean}_{counter}"
            counter += 1

    # Validate or fallback style name
    style_name = str(style).strip() if style else "TableStyleMedium9"
    if not style_name.startswith("TableStyle"):
        if "light" in style_name.lower():
            style_name = "TableStyleLight1"
        elif "dark" in style_name.lower():
            style_name = "TableStyleDark1"
        else:
            style_name = "TableStyleMedium9"

    table = Table(displayName=table_name, ref=ref)
    table_style = TableStyleInfo(
        name=style_name,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=show_row_stripes,
        showColumnStripes=show_column_stripes,
    )
    table.tableStyleInfo = table_style
    ws.add_table(table)
    return table_name


def create_chart(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    chart_type: str = "bar",
    title: str | None = None,
    values_column: Any = None,
    categories_column: Any = None,
    position: str | None = None,
    aggregation: str = "sum",
) -> str:
    """
    Creates an embedded native Excel chart (BarChart, LineChart, PieChart, AreaChart).
    Supports automatic aggregation for large datasets (e.g. Salary by Department, Status counts).
    Returns the cell position where the chart was placed.
    """
    max_col = ws.max_column or 1
    max_row = ws.max_row or 1

    val_col_idx = None
    val_header = None
    if values_column is not None and str(values_column).strip() != "":
        val_col_idx = resolve_column_index(ws, values_column)
        val_header = str(ws.cell(row=1, column=val_col_idx).value or "Value")

    cat_col_idx = None
    cat_header = None
    if categories_column is not None and str(categories_column).strip() != "":
        cat_col_idx = resolve_column_index(ws, categories_column)
        cat_header = str(ws.cell(row=1, column=cat_col_idx).value or "Category")

    # If neither specified, auto-detect numeric and text columns
    if val_col_idx is None and cat_col_idx is None:
        for c in range(1, max_col + 1):
            val = ws.cell(row=2, column=c).value
            if isinstance(val, (int, float)) and val_col_idx is None:
                val_col_idx = c
                val_header = str(ws.cell(row=1, column=c).value or "Value")
            elif isinstance(val, str) and cat_col_idx is None:
                cat_col_idx = c
                cat_header = str(ws.cell(row=1, column=c).value or "Category")

    do_aggregate = False
    if cat_col_idx is not None and max_row > 30:
        do_aggregate = True
    elif val_col_idx is None and cat_col_idx is not None:
        do_aggregate = True

    if do_aggregate and cat_col_idx is not None:
        summary: dict[str, Any] = {}
        for r in range(2, max_row + 1):
            c_val = ws.cell(row=r, column=cat_col_idx).value
            if c_val is None:
                continue
            cat_str = str(c_val).strip()
            if val_col_idx is not None:
                raw_num = ws.cell(row=r, column=val_col_idx).value
                try:
                    num = float(raw_num)
                    summary.setdefault(cat_str, []).append(num)
                except (ValueError, TypeError):
                    pass
            else:
                summary[cat_str] = summary.get(cat_str, 0) + 1

        if val_col_idx is not None:
            if "avg" in aggregation.lower() or "mean" in aggregation.lower():
                agg_data = [(k, sum(v) / len(v)) for k, v in summary.items() if v]
                metric_name = f"Average {val_header}"
            else:
                agg_data = [(k, sum(v)) for k, v in summary.items() if v]
                metric_name = f"Total {val_header}"
        else:
            agg_data = list(summary.items())
            metric_name = f"Count of {cat_header}"

        if len(agg_data) > 25:
            agg_data.sort(key=lambda x: x[1], reverse=True)
            agg_data = agg_data[:25]

        # Write summary table into helper columns
        sum_col = (ws.max_column or 1) + 2
        ws.cell(row=1, column=sum_col, value=cat_header)
        ws.cell(row=1, column=sum_col + 1, value=metric_name)
        for idx, (label, val) in enumerate(agg_data, start=2):
            ws.cell(row=idx, column=sum_col, value=label)
            ws.cell(row=idx, column=sum_col + 1, value=round(val, 2) if isinstance(val, float) else val)

        total_rows = len(agg_data) + 1
        data_ref = Reference(ws, min_col=sum_col + 1, min_row=1, max_row=total_rows)
        cat_ref = Reference(ws, min_col=sum_col, min_row=2, max_row=total_rows)
        default_pos = f"{get_column_letter(sum_col + 3)}2"
    else:
        effective_val_col = val_col_idx or 1
        val_header = val_header or str(ws.cell(row=1, column=effective_val_col).value or "Value")
        limit_row = min(max_row, 31)
        data_ref = Reference(ws, min_col=effective_val_col, min_row=1, max_row=limit_row)
        cat_ref = Reference(ws, min_col=cat_col_idx, min_row=2, max_row=limit_row) if cat_col_idx else None
        default_pos = f"{get_column_letter(max_col + 2)}2"

    c_type = str(chart_type).strip().lower()
    if "pie" in c_type:
        chart = PieChart()
        chart.style = 10
    elif "line" in c_type:
        chart = LineChart()
        chart.style = 13
        if val_header:
            chart.y_axis.title = val_header
        if cat_header:
            chart.x_axis.title = cat_header
    elif "area" in c_type:
        chart = AreaChart()
        chart.style = 11
        if val_header:
            chart.y_axis.title = val_header
        if cat_header:
            chart.x_axis.title = cat_header
    else:
        chart = BarChart()
        chart.type = "col" if "horizontal" not in c_type else "bar"
        chart.style = 10
        if val_header:
            chart.y_axis.title = val_header
        if cat_header:
            chart.x_axis.title = cat_header

    chart.title = title or (f"{val_header} by {cat_header}" if (val_header and cat_header) else (val_header or "Summary Chart"))
    chart.width = 16
    chart.height = 10

    chart.add_data(data_ref, titles_from_data=True)
    if cat_ref is not None:
        chart.set_categories(cat_ref)

    target_pos = position if position else default_pos
    ws.add_chart(chart, target_pos)
    return target_pos


def copy_rows_to_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    target_sheet: str,
    column: Any = None,
    value: Any = None,
    values: list[Any] | None = None,
    operator: str | None = None,
    move: bool = False,
    include_headers: bool = True,
    wb: openpyxl.Workbook | None = None,
) -> tuple[str, int]:
    """
    Finds rows in ws matching the given condition (or all data rows if no condition),
    and copies (or moves) their complete rows with all values, styles, and dimensions into target_sheet.
    If target_sheet does not exist in the workbook, it is created automatically.
    Returns (target_sheet_name, count_of_copied_rows).
    """
    workbook = wb or getattr(ws, "parent", None)
    if workbook is None:
        raise ValueError("Worksheet has no parent workbook.")

    sheet_title = str(target_sheet).strip()
    if not sheet_title:
        sheet_title = "Filtered Data"

    if sheet_title in workbook.sheetnames:
        target_ws = workbook[sheet_title]
        dest_start_row = (target_ws.max_row or 0) + 1
    else:
        target_ws = workbook.create_sheet(title=sheet_title)
        dest_start_row = 1

    max_c = ws.max_column or 1
    max_r = ws.max_row or 1

    # Copy header row if target sheet is newly created or empty
    if include_headers and ((target_ws.max_row or 0) <= 1 and target_ws.cell(row=1, column=1).value is None):
        for col_idx in range(1, max_c + 1):
            src_cell = ws.cell(row=1, column=col_idx)
            dst_cell = target_ws.cell(row=1, column=col_idx)
            dst_cell.value = src_cell.value
            if src_cell.font:
                dst_cell.font = copy(src_cell.font)
            if src_cell.fill:
                dst_cell.fill = copy(src_cell.fill)
            if src_cell.alignment:
                dst_cell.alignment = copy(src_cell.alignment)
            if src_cell.border:
                dst_cell.border = copy(src_cell.border)
            dst_cell.number_format = src_cell.number_format

            col_letter = get_column_letter(col_idx)
            if col_letter in ws.column_dimensions and ws.column_dimensions[col_letter].width:
                target_ws.column_dimensions[col_letter].width = ws.column_dimensions[col_letter].width
        dest_start_row = 2

    # Resolve filter condition
    col_idx = None
    if column is not None and str(column).strip() != "":
        col_idx = resolve_column_index(ws, column)

    # Build target conditions
    parsed_conditions = None
    if value is not None or values is not None:
        raw_targets = values if (values is not None and isinstance(values, list)) else ([value] if not isinstance(value, list) else value)
        parsed_conditions = [parse_operator_and_value(t, operator) for t in raw_targets if t is not None]

    matched_rows = []
    for r in range(2, max_r + 1):
        if parsed_conditions is None:
            matched_rows.append(r)
        elif col_idx is not None:
            cell_val = ws.cell(row=r, column=col_idx).value
            if cell_val is not None:
                for op_sym, target in parsed_conditions:
                    if matches_condition(cell_val, op_sym, target):
                        matched_rows.append(r)
                        break
        else:
            row_matched = False
            for c in range(1, max_c + 1):
                c_val = ws.cell(row=r, column=c).value
                if c_val is not None:
                    for op_sym, target in parsed_conditions:
                        if matches_condition(c_val, op_sym, target):
                            matched_rows.append(r)
                            row_matched = True
                            break
                if row_matched:
                    break

    # Copy matched rows
    curr_dest_r = dest_start_row
    for src_r in matched_rows:
        for c in range(1, max_c + 1):
            src_cell = ws.cell(row=src_r, column=c)
            dst_cell = target_ws.cell(row=curr_dest_r, column=c)
            dst_cell.value = src_cell.value
            if src_cell.font:
                dst_cell.font = copy(src_cell.font)
            if src_cell.fill:
                dst_cell.fill = copy(src_cell.fill)
            if src_cell.alignment:
                dst_cell.alignment = copy(src_cell.alignment)
            if src_cell.border:
                dst_cell.border = copy(src_cell.border)
            dst_cell.number_format = src_cell.number_format
        curr_dest_r += 1

    # If move, delete from source in reverse
    if move and matched_rows:
        for src_r in reversed(matched_rows):
            ws.delete_rows(src_r)

    workbook.active = target_ws
    return sheet_title, len(matched_rows)


def create_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    sheet_name: str,
    headers: list[Any] | None = None,
    position: Any = None,
    wb: openpyxl.Workbook | None = None,
) -> str:
    """
    Creates a new sheet in the workbook with optional headers in row 1.
    Returns the created sheet title.
    """
    workbook = wb or getattr(ws, "parent", None)
    if workbook is None:
        raise ValueError("Worksheet has no parent workbook.")

    title = str(sheet_name).strip()
    if not title:
        title = f"Sheet_{len(workbook.sheetnames) + 1}"

    index = None
    if position is not None:
        pos_str = str(position).strip().lower()
        if pos_str in ("first", "start", "0", "1"):
            index = 0
        elif pos_str in ("last", "end"):
            index = len(workbook.sheetnames)
        elif pos_str.isdigit():
            index = max(0, min(int(pos_str) - 1, len(workbook.sheetnames)))

    new_ws = workbook.create_sheet(title=title, index=index)

    if headers and isinstance(headers, list):
        for col_idx, h in enumerate(headers, start=1):
            new_ws.cell(row=1, column=col_idx, value=str(h))

    workbook.active = new_ws
    return new_ws.title


def duplicate_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    new_sheet_name: str | None = None,
    source_sheet: str | None = None,
    wb: openpyxl.Workbook | None = None,
) -> str:
    """
    Duplicates an entire sheet with all data, styles, and formulas.
    Returns the new sheet title.
    """
    workbook = wb or getattr(ws, "parent", None)
    if workbook is None:
        raise ValueError("Worksheet has no parent workbook.")

    src_ws = workbook[source_sheet] if source_sheet and source_sheet in workbook.sheetnames else ws
    title = str(new_sheet_name).strip() if new_sheet_name else f"{src_ws.title}_Copy"

    target_ws = workbook.copy_worksheet(src_ws)
    target_ws.title = title
    workbook.active = target_ws
    return target_ws.title


def rename_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    new_name: str,
    old_name: str | None = None,
    wb: openpyxl.Workbook | None = None,
) -> str:
    """
    Renames a sheet in the workbook.
    """
    workbook = wb or getattr(ws, "parent", None)
    if workbook is None:
        raise ValueError("Worksheet has no parent workbook.")

    target_ws = workbook[old_name] if old_name and old_name in workbook.sheetnames else ws
    target_ws.title = str(new_name).strip()
    return target_ws.title


def delete_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    sheet_name: str,
    wb: openpyxl.Workbook | None = None,
) -> str:
    """
    Deletes a sheet by name from the workbook.
    """
    workbook = wb or getattr(ws, "parent", None)
    if workbook is None:
        raise ValueError("Worksheet has no parent workbook.")
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {workbook.sheetnames}")
    if len(workbook.sheetnames) <= 1:
        raise ValueError("Cannot delete the only sheet in the workbook.")

    to_delete = workbook[sheet_name]
    workbook.remove(to_delete)
    return sheet_name


def export_to_file(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    output_file: str | None = None,
    column: Any = None,
    value: Any = None,
    values: list[Any] | None = None,
    operator: str | None = None,
    sheet_name: str = "Sheet1",
) -> tuple[str, int]:
    """
    Exports filtered rows (or all data rows) into a brand new, separate .xlsx file.
    Returns (output_filepath, count_of_exported_rows).
    """
    if not output_file:
        output_file = "exported_data.xlsx"
    out_str = str(output_file).strip()
    if not out_str.lower().endswith(".xlsx"):
        out_str += ".xlsx"

    new_wb = openpyxl.Workbook()
    dest_ws = new_wb.active
    dest_ws.title = str(sheet_name).strip() or "Sheet1"

    target_title, count = copy_rows_to_sheet(
        ws,
        target_sheet=dest_ws.title,
        column=column,
        value=value,
        values=values,
        operator=operator,
        include_headers=True,
        wb=new_wb,
    )
    new_wb.save(out_str)
    return out_str, count


# =====================================================================
# Action Router & Operation Validator
# =====================================================================

def validate_operation(ws: openpyxl.worksheet.worksheet.Worksheet, operation: dict[str, Any]) -> None:
    """
    Validates that:
      - action is supported
      - row number is valid
      - column exists
      - width is a valid number
    Raises ValueError with a helpful message if invalid.
    """
    if not isinstance(operation, dict):
        raise ValueError(f"Operation must be a JSON object/dict, got: {type(operation).__name__}")

    action = operation.get("action")
    if not action or action not in ALLOWED_ACTIONS:
        raise ValueError(f"Unsupported action: '{action}'. Allowed actions: {sorted(ALLOWED_ACTIONS)}")

    # Action-specific validations
    if action in {"highlight_row", "bold_row", "delete_row"}:
        if "row" not in operation:
            raise ValueError(f"Action '{action}' requires 'row' parameter.")
        resolve_row_number(ws, operation["row"])

    if action in {"highlight_column", "bold_column"}:
        if "column" not in operation:
            raise ValueError(f"Action '{action}' requires 'column' parameter.")
        resolve_column_index(ws, operation["column"])

    if action == "delete_column":
        if "column" not in operation and "columns" not in operation:
            raise ValueError("Action 'delete_column' requires 'column' parameter.")
        cols = operation.get("columns", operation.get("column"))
        if isinstance(cols, list):
            for c in cols:
                resolve_column_index(ws, c)
        elif isinstance(cols, str) and (" and " in cols.lower() or "," in cols):
            split_cols = [c.strip() for c in re.split(r"\s+and\s+|,", cols, flags=re.IGNORECASE) if c.strip()]
            for c in split_cols:
                resolve_column_index(ws, c)
        else:
            resolve_column_index(ws, cols)

    if action == "delete_rows_where":
        if "value" not in operation and "values" not in operation:
            raise ValueError("Action 'delete_rows_where' requires 'value' or 'values' parameter.")
        if "column" in operation and operation["column"]:
            resolve_column_index(ws, operation["column"])

    if action == "bold_cell":
        if "cell" not in operation:
            raise ValueError("Action 'bold_cell' requires 'cell' parameter.")

    if action == "bold_value":
        if "value" not in operation:
            raise ValueError("Action 'bold_value' requires 'value' parameter.")
        if "column" in operation and operation["column"]:
            resolve_column_index(ws, operation["column"])

    if action == "resize_column":
        if "column" not in operation:
            raise ValueError("Action 'resize_column' requires 'column' parameter.")
        if "width" not in operation:
            raise ValueError("Action 'resize_column' requires 'width' parameter.")
        resolve_column_index(ws, operation["column"])
        try:
            w = float(operation["width"])
            if w <= 0:
                raise ValueError()
        except (ValueError, TypeError):
            raise ValueError(f"Width '{operation.get('width')}' is not a valid positive number.")

    if action == "highlight_rows_where":
        if "column" not in operation:
            raise ValueError("Action 'highlight_rows_where' requires 'column' parameter.")
        if "value" not in operation:
            raise ValueError("Action 'highlight_rows_where' requires 'value' parameter.")
        resolve_column_index(ws, operation["column"])

    if action == "highlight_value":
        if "value" not in operation:
            raise ValueError("Action 'highlight_value' requires 'value' parameter.")
        if "column" in operation and operation["column"]:
            resolve_column_index(ws, operation["column"])

    if action == "update_value_where":
        if "search_value" not in operation and "value" in operation:
            operation["search_value"] = operation["value"]
        if "search_value" not in operation:
            raise ValueError("Action 'update_value_where' requires 'search_value' parameter.")
        if "target_column" not in operation and "column" in operation:
            operation["target_column"] = operation["column"]
        if "target_column" not in operation:
            raise ValueError("Action 'update_value_where' requires 'target_column' parameter.")
        if "new_value" not in operation:
            raise ValueError("Action 'update_value_where' requires 'new_value' parameter.")
        resolve_column_index(ws, operation["target_column"])
        if "search_column" in operation and operation["search_column"]:
            resolve_column_index(ws, operation["search_column"])

    if action == "update_cell":
        if "cell" not in operation:
            raise ValueError("Action 'update_cell' requires 'cell' parameter.")
        if "value" not in operation:
            raise ValueError("Action 'update_cell' requires 'value' parameter.")

    if action == "insert_row":
        if "row" in operation and operation["row"] is not None:
            resolve_row_number(ws, operation["row"])
        if "after_row" in operation and operation["after_row"] is not None:
            resolve_row_number(ws, operation["after_row"])

    if action == "insert_column":
        if "after_column" in operation and operation["after_column"] is not None:
            resolve_column_index(ws, operation["after_column"])

    if action in {"sort_rows", "arrange_rows"}:
        col = operation.get("by_column", operation.get("column"))
        if not col:
            raise ValueError(f"Action '{action}' requires 'by_column' parameter.")
        operation["by_column"] = col
        resolve_column_index(ws, col)

    if action in {"sort_columns", "arrange_columns"}:
        pass

    if action == "reorder_columns":
        if "order" not in operation and "columns" not in operation and "move_column" not in operation:
            raise ValueError("Action 'reorder_columns' requires 'columns' or 'move_column' parameter.")

    if action == "create_table":
        pass

    if action == "create_chart":
        if "values_column" in operation and operation["values_column"]:
            resolve_column_index(ws, operation["values_column"])
        if "categories_column" in operation and operation["categories_column"]:
            resolve_column_index(ws, operation["categories_column"])

    if action in {"copy_rows_to_sheet", "move_rows_to_sheet"}:
        target = operation.get("target_sheet", operation.get("sheet_name", operation.get("sheet")))
        if not target:
            raise ValueError(f"Action '{action}' requires 'target_sheet' parameter.")
        operation["target_sheet"] = target
        if "column" in operation and operation["column"]:
            resolve_column_index(ws, operation["column"])

    if action == "create_sheet":
        if "sheet_name" not in operation and "name" not in operation:
            raise ValueError("Action 'create_sheet' requires 'sheet_name' parameter.")

    if action == "duplicate_sheet":
        pass

    if action == "rename_sheet":
        if "new_name" not in operation and "name" not in operation:
            raise ValueError("Action 'rename_sheet' requires 'new_name' parameter.")

    if action == "delete_sheet":
        if "sheet_name" not in operation and "name" not in operation:
            raise ValueError("Action 'delete_sheet' requires 'sheet_name' parameter.")

    if action == "export_to_file":
        if "output_file" not in operation and "file" not in operation and "filename" not in operation:
            raise ValueError("Action 'export_to_file' requires 'output_file' parameter.")


def execute_operation(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    operation: dict[str, Any],
    wb: openpyxl.Workbook | None = None,
) -> bool:
    """
    Action Router:
    Inspects operation['action'] and calls the appropriate predefined function.
    Validates before execution. Does NOT modify workbook if validation fails.
    Returns True on success, False on failure.
    """
    action = operation.get("action")
    workbook = wb or getattr(ws, "parent", None)
    print(f"\nExecuting operation: {action}")

    try:
        validate_operation(ws, operation)
    except ValueError as e:
        print(f"Validation Error: {e}")
        return False

    try:
        if action == "highlight_row":
            row = operation["row"]
            color = operation.get("color", "FFFF00")
            highlight_row(ws, row, color)
            print(f"  -> Successfully highlighted row {row} with color {standardize_color(color)}")

        elif action == "highlight_column":
            col = operation["column"]
            color = operation.get("color", "FFFF00")
            highlight_column(ws, col, color)
            print(f"  -> Successfully highlighted column '{col}' with color {standardize_color(color)}")

        elif action == "highlight_rows_where":
            col = operation["column"]
            val = operation["value"]
            color = operation.get("color", "FFFF00")
            op_sym = operation.get("operator")
            rows = highlight_rows_where(ws, col, val, color, operator=op_sym)
            print(f"  -> Successfully highlighted {len(rows)} row(s) ({rows[:10]}{'...' if len(rows) > 10 else ''}) where '{col}' satisfies '{val}' with color {standardize_color(color)}")

        elif action == "highlight_value":
            val = operation["value"]
            col = operation.get("column")
            color = operation.get("color", "FFFF00")
            cells = highlight_value(ws, val, col, color)
            print(f"  -> Successfully highlighted {len(cells)} cell(s) ({cells[:10]}{'...' if len(cells) > 10 else ''}) containing '{val}' with color {standardize_color(color)}")

        elif action == "bold_row":
            row = operation["row"]
            bold_row(ws, row)
            print(f"  -> Successfully made row {row} bold")

        elif action == "bold_column":
            col = operation["column"]
            bold_column(ws, col)
            print(f"  -> Successfully made column '{col}' bold")

        elif action == "bold_cell":
            cell_coord = operation["cell"]
            bold_cell(ws, cell_coord)
            print(f"  -> Successfully made cell {cell_coord} bold")

        elif action == "bold_value":
            val = operation["value"]
            col = operation.get("column")
            coords = bold_value(ws, val, col)
            print(f"  -> Successfully made cell(s) {coords} containing '{val}' bold")

        elif action == "delete_row":
            row = operation["row"]
            delete_row(ws, row)
            print(f"  -> Successfully deleted row {row}")

        elif action == "delete_rows_where":
            col = operation.get("column")
            val = operation.get("value")
            vals = operation.get("values")
            op_sym = operation.get("operator")
            deleted = delete_rows_where(ws, column=col, value=val, values=vals, operator=op_sym)
            col_name = f"column '{col}'" if col else "any column"
            print(f"  -> Successfully deleted {len(deleted)} row(s) where {col_name} satisfies '{val or vals}'")

        elif action == "delete_column":
            col = operation.get("column", operation.get("columns"))
            del_indices = delete_column(ws, col)
            print(f"  -> Successfully deleted {len(del_indices)} column(s) ({col})")

        elif action == "resize_column":
            col = operation["column"]
            width = operation["width"]
            resize_column(ws, col, width)
            print(f"  -> Successfully resized column '{col}' to width {width}")

        elif action == "insert_row":
            target_row = insert_row(
                ws,
                row_number=operation.get("row"),
                after_row=operation.get("after_row"),
                values=operation.get("values"),
            )
            print(f"  -> Successfully inserted row at index {target_row}")

        elif action == "append_row":
            target_row = append_row(ws, values=operation.get("values"))
            print(f"  -> Successfully appended row at index {target_row}")

        elif action == "insert_column":
            target_col = insert_column(
                ws,
                column=operation.get("column"),
                after_column=operation.get("after_column"),
                header=operation.get("header"),
                values=operation.get("values"),
            )
            print(f"  -> Successfully inserted column at index {target_col}")

        elif action == "update_value_where":
            search_val = operation["search_value"]
            search_col = operation.get("search_column")
            target_col = operation["target_column"]
            new_val = operation["new_value"]
            updated = update_value_where(ws, search_val, target_col, new_val, search_col)
            print(f"  -> Successfully updated cell(s) {updated} in column '{target_col}' to '{new_val}' where matching '{search_val}'")

        elif action == "update_cell":
            cell_coord = operation["cell"]
            new_val = operation["value"]
            res = update_cell(ws, cell_coord, new_val)
            print(f"  -> Successfully updated cell {res} to '{new_val}'")

        elif action in {"sort_rows", "arrange_rows"}:
            col = operation.get("by_column", operation.get("column"))
            order = operation.get("order", "ascending")
            count = sort_rows(ws, by_column=col, order=order)
            print(f"  -> Successfully arranged {count} row(s) in {order} order of '{col}'")

        elif action in {"sort_columns", "arrange_columns"}:
            order = operation.get("order", "ascending")
            cols = sort_columns(ws, order=order)
            print(f"  -> Successfully arranged columns in {order} order: {cols}")

        elif action == "reorder_columns":
            cols_order = operation.get("columns", operation.get("order"))
            move_col = operation.get("move_column")
            before_col = operation.get("before_column")
            after_col = operation.get("after_column")
            pos = operation.get("position")
            new_cols = reorder_columns(
                ws,
                order=cols_order,
                move_column=move_col,
                before_column=before_col,
                after_column=after_col,
                position=pos,
            )
            print(f"  -> Successfully reordered columns to: {new_cols}")

        elif action == "create_table":
            name = operation.get("name")
            ref = operation.get("ref")
            style = operation.get("style", "TableStyleMedium9")
            stripes = operation.get("show_row_stripes", True)
            t_name = create_table(ws, name=name, ref=ref, style=style, show_row_stripes=stripes)
            print(f"  -> Successfully created Excel Table '{t_name}' with style '{style}'")

        elif action == "create_chart":
            c_type = operation.get("chart_type", "bar")
            title = operation.get("title")
            val_col = operation.get("values_column")
            cat_col = operation.get("categories_column")
            pos = operation.get("position")
            agg = operation.get("aggregation", "sum")
            placed_at = create_chart(
                ws,
                chart_type=c_type,
                title=title,
                values_column=val_col,
                categories_column=cat_col,
                position=pos,
                aggregation=agg,
            )
            print(f"  -> Successfully created {c_type} chart '{title or 'Chart'}' placed at {placed_at}")

        elif action in {"copy_rows_to_sheet", "move_rows_to_sheet"}:
            target_sheet = operation.get("target_sheet", operation.get("sheet_name", operation.get("sheet")))
            col = operation.get("column")
            val = operation.get("value")
            vals = operation.get("values")
            op_sym = operation.get("operator")
            is_move = action == "move_rows_to_sheet" or operation.get("move", False)
            sheet_title, count = copy_rows_to_sheet(
                ws,
                target_sheet=target_sheet,
                column=col,
                value=val,
                values=vals,
                operator=op_sym,
                move=is_move,
                include_headers=operation.get("include_headers", True),
                wb=workbook,
            )
            verb = "moved" if is_move else "copied"
            col_desc = f"where '{col}' satisfies '{val or vals}'" if (val or vals) else "from active sheet"
            print(f"  -> Successfully {verb} {count} row(s) {col_desc} into sheet '{sheet_title}'")

        elif action == "create_sheet":
            s_name = operation.get("sheet_name", operation.get("name"))
            hdrs = operation.get("headers")
            pos = operation.get("position")
            created_title = create_sheet(ws, sheet_name=s_name, headers=hdrs, position=pos, wb=workbook)
            print(f"  -> Successfully created new sheet '{created_title}'")

        elif action == "duplicate_sheet":
            new_name = operation.get("new_sheet_name", operation.get("name"))
            src_name = operation.get("source_sheet")
            dup_title = duplicate_sheet(ws, new_sheet_name=new_name, source_sheet=src_name, wb=workbook)
            print(f"  -> Successfully duplicated sheet to '{dup_title}'")

        elif action == "rename_sheet":
            new_name = operation.get("new_name", operation.get("name"))
            old_name = operation.get("old_name")
            renamed = rename_sheet(ws, new_name=new_name, old_name=old_name, wb=workbook)
            print(f"  -> Successfully renamed sheet to '{renamed}'")

        elif action == "delete_sheet":
            s_name = operation.get("sheet_name", operation.get("name"))
            deleted = delete_sheet(ws, sheet_name=s_name, wb=workbook)
            print(f"  -> Successfully deleted sheet '{deleted}'")

        elif action == "export_to_file":
            out_file = operation.get("output_file", operation.get("file", operation.get("filename")))
            col = operation.get("column")
            val = operation.get("value")
            vals = operation.get("values")
            op_sym = operation.get("operator")
            s_name = operation.get("sheet_name", "Sheet1")
            file_path, count = export_to_file(
                ws,
                output_file=out_file,
                column=col,
                value=val,
                values=vals,
                operator=op_sym,
                sheet_name=s_name,
            )
            print(f"  -> Successfully exported {count} row(s) into new file: {file_path}")

        return True

    except Exception as e:
        print(f"Execution Error executing '{action}': {e}")
        return False

