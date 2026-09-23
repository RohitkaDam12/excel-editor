"""
llm.py - Isolated LLM client and instruction parser supporting Google Gemini.

This module communicates with the Gemini API (via google-genai) or OpenAI API
to convert natural-language user requests into a strictly structured JSON operations payload.
It never generates or executes Python code.
"""

import json
import os
import re
from typing import Any
from dotenv import load_dotenv

from prompts import SYSTEM_PROMPT, build_user_message

# Load environment variables from .env file if present
load_dotenv()


def _clean_and_parse_json(text: str) -> dict[str, Any]:
    """
    Cleans markdown code fences and parses raw LLM output into a dictionary.
    """
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)

    parsed = json.loads(text.strip())
    if not isinstance(parsed, dict) or "operations" not in parsed:
        raise ValueError("LLM response must be a JSON object containing an 'operations' array.")
    return parsed


def _parse_single_clause(clause: str) -> dict[str, Any] | None:
    """
    Parses a single atomic instruction clause into an operation dictionary.
    """
    clause = clause.strip().lower()
    if not clause:
        return None

    # Sort / Arrange Rows (e.g., "arrange the rows according to ascending order of age", "sort rows by salary descending")
    if any(k in clause for k in ["sort", "arrange", "order"]) and any(k in clause for k in ["row", "rows"]):
        order = "ascending"
        if any(k in clause for k in ["desc", "reverse", "z to a", "z-a", "high to low", "largest", "descending"]):
            order = "descending"

        by_col = "1"
        m_col = re.search(r"(?:by|of|on|according to|based on)\s+(?:ascending\s+order\s+of\s+|descending\s+order\s+of\s+)?([a-zA-Z0-9_\s]+)", clause)
        if m_col:
            raw_col = m_col.group(1).strip()
            raw_col = re.sub(r"\s+(?:ascending|descending|asc|desc|in\s+excel|from\s+excel).*$", "", raw_col, flags=re.IGNORECASE).strip()
            raw_col = re.sub(r"^(?:order\s+of\s+|ascending\s+order\s+of\s+|descending\s+order\s+of\s+)", "", raw_col, flags=re.IGNORECASE).strip()
            if raw_col:
                by_col = raw_col.capitalize()
        return {"action": "sort_rows", "by_column": by_col, "order": order}

    # Sort / Arrange Columns (e.g., "arrange columns alphabetically", "sort columns A to Z")
    if any(k in clause for k in ["sort", "arrange", "order"]) and ("column" in clause or "columns" in clause or "col" in clause):
        order = "ascending"
        if any(k in clause for k in ["desc", "reverse", "z to a", "z-a"]):
            order = "descending"
        return {"action": "sort_columns", "order": order}

    # Move Column (e.g., "move column City after Country")
    if "move" in clause and ("column" in clause or "col" in clause):
        m_move = re.search(r"move\s+(?:column\s+)?([a-zA-Z0-9_]+)\s+(after|before|to)\s+(.+)", clause)
        if m_move:
            col_target = m_move.group(1).strip().capitalize()
            rel = m_move.group(2).strip().lower()
            dest = m_move.group(3).strip()
            dest_clean = re.sub(r"^(?:the\s+|column\s+)", "", dest, flags=re.IGNORECASE).strip()
            op = {"action": "reorder_columns", "move_column": col_target}
            if rel == "after":
                op["after_column"] = dest_clean.capitalize()
            elif rel == "before":
                op["before_column"] = dest_clean.capitalize()
            elif rel == "to":
                op["position"] = dest_clean.lower()
            return op

    # Create Table
    if ("table" in clause and any(k in clause for k in ["create", "make", "convert", "format", "turn", "add"])) and "chart" not in clause:
        return {"action": "create_table"}

    # Create Chart
    if any(k in clause for k in ["chart", "graph", "plot"]):
        c_type = "bar"
        if "pie" in clause:
            c_type = "pie"
        elif "line" in clause:
            c_type = "line"
        elif "area" in clause:
            c_type = "area"

        m_of_by = re.search(r"of\s+([a-zA-Z0-9_\s]+?)\s+by\s+([a-zA-Z0-9_\s]+)", clause)
        if m_of_by:
            v_col = m_of_by.group(1).strip().capitalize()
            c_col = m_of_by.group(2).strip().capitalize()
            return {
                "action": "create_chart",
                "chart_type": c_type,
                "values_column": v_col,
                "categories_column": c_col,
                "title": f"{v_col} by {c_col}",
            }

        m_of = re.search(r"of\s+([a-zA-Z0-9_\s]+)", clause)
        if m_of:
            target_col = m_of.group(1).strip()
            target_col = re.sub(r"\s+(?:in\s+excel|from\s+excel).*$", "", target_col, flags=re.IGNORECASE).strip().capitalize()
            if c_type == "pie":
                return {
                    "action": "create_chart",
                    "chart_type": "pie",
                    "categories_column": target_col,
                    "title": f"{target_col} Breakdown",
                }
            return {
                "action": "create_chart",
                "chart_type": c_type,
                "values_column": target_col,
                "title": f"{target_col} Chart",
            }

        return {"action": "create_chart", "chart_type": c_type}

    # Copy / Move rows to sheet (e.g. "Find patients from the Cardiology department and copy their complete rows into a new sheet called Cardiology Patients")
    if ("sheet" in clause) and any(k in clause for k in ["copy", "move", "find", "filter", "transfer", "put", "send"]):
        sheet_target = "Filtered Data"
        m_sname = re.search(r"(?:sheet\s+(?:called|named)\s+|sheet\s+to\s+|into\s+a\s+new\s+sheet\s+|into\s+sheet\s+|to\s+sheet\s+)(['\"]?[a-zA-Z0-9_\s]+['\"]?)", clause)
        if m_sname:
            sheet_target = m_sname.group(1).strip().strip("'\"").title()

        col_found = None
        val_found = None
        m_dep = re.search(r"(?:from|in|where)\s+(?:the\s+)?([a-zA-Z0-9_]+)\s+([a-zA-Z0-9_]+)", clause)
        if m_dep:
            w1, w2 = m_dep.group(1).strip(), m_dep.group(2).strip()
            if any(h in w2.lower() for h in ["department", "dept", "city", "status", "company", "role", "title"]):
                val_found = w1.capitalize()
                col_found = w2.capitalize()
            elif any(h in w1.lower() for h in ["department", "dept", "city", "status", "company", "role", "title"]):
                col_found = w1.capitalize()
                val_found = w2.capitalize()

        m_where = re.search(r"(?:where|with)\s+([a-zA-Z0-9_\s]+?)\s*(?:is|=|==|<|>|<=|>=)?\s*([a-zA-Z0-9_\-<>=\s]+)", clause)
        if m_where and not col_found:
            col_found = m_where.group(1).strip().capitalize()
            val_found = m_where.group(2).strip()

        is_move = "move" in clause or "transfer" in clause
        op = {
            "action": "copy_rows_to_sheet",
            "target_sheet": sheet_target,
            "move": is_move,
        }
        if col_found:
            op["column"] = col_found
        if val_found:
            op["value"] = val_found
        return op

    # Export to new Excel file
    if any(k in clause for k in ["export", "save"]) and any(k in clause for k in ["file", "workbook", "excel file", ".xlsx"]):
        out_f = "exported_data.xlsx"
        m_file = re.search(r"([a-zA-Z0-9_\-]+\.xlsx)", clause)
        if m_file:
            out_f = m_file.group(1).strip()
        else:
            m_called = re.search(r"(?:called|named)\s+([a-zA-Z0-9_\-]+)", clause)
            if m_called:
                out_f = f"{m_called.group(1).strip()}.xlsx"

        op = {"action": "export_to_file", "output_file": out_f}
        m_where = re.search(r"(?:where|with)\s+([a-zA-Z0-9_\s]+?)\s*(?:is|=|==|<|>|<=|>=)?\s*([a-zA-Z0-9_\-<>=\s]+)", clause)
        if m_where:
            op["column"] = m_where.group(1).strip().capitalize()
            op["value"] = m_where.group(2).strip()
        return op

    # Create empty sheet
    if "sheet" in clause and any(k in clause for k in ["create", "add", "new sheet"]):
        s_title = "Sheet"
        m_s = re.search(r"sheet\s+(?:called|named)\s+([a-zA-Z0-9_\s]+)", clause)
        if m_s:
            s_title = m_s.group(1).strip().title()
        return {"action": "create_sheet", "sheet_name": s_title}

    # Highlight Row
    if "highlight" in clause and "row" in clause:
        row = 1
        m = re.search(r"row\s*(\d+)", clause)
        if m:
            row = int(m.group(1))
        color = "FFFF00"
        if "red" in clause:
            color = "FF0000"
        elif "green" in clause:
            color = "00FF00"
        elif "blue" in clause:
            color = "0000FF"
        return {"action": "highlight_row", "row": row, "color": color}

    # Bold Row
    if "bold" in clause and "row" in clause:
        row = 1
        m = re.search(r"row\s*(\d+)", clause)
        if m:
            row = int(m.group(1))
        return {"action": "bold_row", "row": row}

    # Bold Column
    if "bold" in clause and ("column" in clause or "col" in clause):
        col = "1"
        if "last" in clause:
            col = "last"
        elif "email" in clause:
            col = "Email"
        elif "name" in clause:
            col = "Name"
        else:
            m = re.search(r"column\s*([a-zA-Z0-9_]+)", clause)
            if m:
                col = m.group(1)
        return {"action": "bold_column", "column": col}

    # Bold Value / Name (e.g. "bold the name Vikram joshi" or "bold Emily Brown")
    if "bold" in clause:
        # Check if column is specified
        target_col = None
        if "name" in clause:
            target_col = "name"
        elif "company" in clause:
            target_col = "company"
        elif "email" in clause:
            target_col = "email"

        val_match = re.search(r"bold\s+(?:the\s+)?(?:name\s+|cell\s+|value\s+)?(.+)", clause, flags=re.IGNORECASE)
        val = val_match.group(1).strip() if val_match else clause.replace("bold", "").strip()
        # Clean up filler words
        val = re.sub(r"^(?:the\s+|name\s+|cell\s+|value\s+)", "", val, flags=re.IGNORECASE).strip()
        if val:
            op = {"action": "bold_value", "value": val}
            if target_col:
                op["column"] = target_col
            return op

    # Insert / Append Row
    if any(k in clause for k in ["add", "insert"]) and "row" in clause:
        if any(k in clause for k in ["bottom", "end", "last"]):
            return {"action": "append_row"}
        m_after = re.search(r"after\s+(?:row\s+)?(\d+)", clause)
        if m_after:
            return {"action": "insert_row", "after_row": int(m_after.group(1))}
        m_row = re.search(r"row\s*(\d+)", clause)
        row_num = int(m_row.group(1)) if m_row else None
        return {"action": "insert_row", "row": row_num}

    # Insert Column
    if any(k in clause for k in ["add", "insert"]) and ("col" in clause or "column" in clause):
        header_name = "New Column"
        m_head = re.search(r"(?:called|named|name)\s+([a-zA-Z0-9_]+)", clause)
        if m_head:
            header_name = m_head.group(1)

        after_col = None
        m_after = re.search(r"(?:after|next to|right of|right side of)\s+([a-zA-Z0-9_]+)", clause)
        if m_after:
            after_col = m_after.group(1)

        op = {"action": "insert_column", "header": header_name}
        if after_col:
            op["after_column"] = after_col
        elif any(k in clause for k in ["end", "last"]):
            op["column"] = "last"
        return op

    # Delete Column
    if any(k in clause for k in ["delete", "remove", "drop"]) and ("column" in clause or "col" in clause):
        col = "last"
        if "first" in clause:
            col = "1"
        elif "last" in clause:
            col = "last"
        elif "email" in clause:
            col = "Email"
        else:
            m_named = re.search(r"col(?:umn)?\s+(?:named\s+|called\s+)?([a-zA-Z0-9_]+)", clause)
            m_pre = re.search(r"(?:the\s+)?([a-zA-Z0-9_]+)\s+col(?:umn)?", clause)
            if m_named and m_named.group(1).lower() not in ("the", "a", "an", "last", "first"):
                col = m_named.group(1)
            elif m_pre and m_pre.group(1).lower() not in ("delete", "remove", "drop", "the", "a", "an", "last", "first"):
                col = m_pre.group(1)
        return {"action": "delete_column", "column": col}

    # Delete Rows (Conditional or specific)
    if any(k in clause for k in ["delete", "remove", "drop"]) and any(k in clause for k in ["row", "rows", "record", "records"]):
        # 1. Check for named/value row deletion (e.g. "delete the row named rohit and vikram", "delete row named rohit")
        m_named = re.search(r"(?:named|name|called)\s+(.+)", clause)
        if m_named:
            names_raw = m_named.group(1).strip()
            names_raw = re.sub(r"\s+from\s+.*$", "", names_raw, flags=re.IGNORECASE).strip()
            names = [n.strip().capitalize() for n in re.split(r"\s+and\s+|,", names_raw) if n.strip()]
            if len(names) > 1:
                return {"action": "delete_rows_where", "values": names}
            return {"action": "delete_rows_where", "value": names[0]}

        # 2. Check for condition: "where ...", "with ...", "having ..."
        m_where = re.search(r"(?:where|with|having)\s+([a-zA-Z0-9_\s]+?)\s*(?:is\s+|==\s*|=\s*|:\s*)?(<.*|>.*|<=.*|>=.*|!=.*|[a-zA-Z0-9_\-]+)", clause)
        if m_where:
            col_part = m_where.group(1).strip()
            val_part = m_where.group(2).strip()
            return {"action": "delete_rows_where", "column": col_part, "value": val_part}

        # 3. Check for pattern: "delete the pune city rows" / "delete pune rows"
        m_city_rows = re.search(r"(?:the\s+)?([a-zA-Z0-9_]+)\s+([a-zA-Z0-9_]+)\s+rows?", clause)
        if m_city_rows and m_city_rows.group(1).lower() not in ("delete", "remove", "drop", "the", "all", "first", "last"):
            val_part = m_city_rows.group(1).strip()
            col_part = m_city_rows.group(2).strip()
            return {"action": "delete_rows_where", "column": col_part.capitalize(), "value": val_part.capitalize()}

        # 4. Check for direct row number: "delete row 5", "delete the first row", "delete last row"
        row = 1
        if "last" in clause:
            row = "last"
        else:
            m = re.search(r"row\s*(\d+)", clause)
            if m:
                row = int(m.group(1))
            elif "first" in clause:
                row = 1
        return {"action": "delete_row", "row": row}

    # Resize Column
    if any(k in clause for k in ["resize", "width", "wider"]) and any(k in clause for k in ["col", "column", "b", "email"]):
        col = "B"
        if "email" in clause:
            col = "Email"
        elif "last" in clause:
            col = "last"
        else:
            m_pre = re.search(r"(?:the\s+)?([a-zA-Z0-9_]+)\s+col", clause)
            m_post = re.search(r"col(?:umn)?\s+([a-zA-Z0-9_]+)", clause)
            if m_pre and m_pre.group(1) not in ("resize", "the", "a", "an", "make"):
                col = m_pre.group(1)
            elif m_post:
                col = m_post.group(1)

        width = 30
        w_match = re.search(r"width\s*(\d+)", clause) or re.search(r"(\d+)\s*width", clause) or re.search(r"to\s*(\d+)", clause)
        if w_match:
            width = int(w_match.group(1))
        elif "wider" in clause:
            width = 25
        return {"action": "resize_column", "column": col, "width": width}

    # Highlight Column
    if "highlight" in clause and ("col" in clause or "column" in clause or any(h in clause for h in ["email", "name", "status", "company", "salary"])):
        col = "1"
        if "last" in clause:
            col = "last"
        elif "email" in clause:
            col = "Email"
        elif "name" in clause:
            col = "Name"
        elif "status" in clause:
            col = "Status"
        elif "company" in clause:
            col = "Company"
        else:
            m_pre = re.search(r"(?:the\s+)?([a-zA-Z0-9_]+)\s+col", clause)
            m_post = re.search(r"col(?:umn)?\s+([a-zA-Z0-9_]+)", clause)
            if m_pre and m_pre.group(1) not in ("highlight", "the", "a", "an"):
                col = m_pre.group(1)
            elif m_post:
                col = m_post.group(1)

        color = "FFFF00"
        if "red" in clause:
            color = "FF0000"
        elif "green" in clause:
            color = "00FF00"
        elif "blue" in clause:
            color = "0000FF"
        return {"action": "highlight_column", "column": col, "color": color}

    return None


def _fallback_parse(user_prompt: str) -> dict[str, Any]:
    """
    Offline fallback parser used when no API key is configured.
    Splits compound instructions (e.g. by 'and', ';', ',') and parses each clause.
    """
    clauses = re.split(r"\s+and\s+|;|,", user_prompt, flags=re.IGNORECASE)
    operations: list[dict[str, Any]] = []

    for clause in clauses:
        op = _parse_single_clause(clause)
        if op:
            operations.append(op)

    if not operations:
        raise ValueError(f"Could not parse instruction offline without an LLM API: '{user_prompt}'")

    return {"operations": operations}


def _call_gemini(api_key: str, user_prompt: str, headers: list[str] | None = None) -> dict[str, Any]:
    """Calls Google Gemini API using google-genai SDK with automatic model fallback."""
    import warnings
    warnings.filterwarnings("ignore")

    try:
        import google.genai.models as _gm
        _gm._response_function_calls_warning_logged = True
    except Exception:
        pass

    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    preferred_model = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite")
    
    candidate_models = [
        preferred_model,
        "gemini-3.5-flash-lite",
        "gemini-flash-latest",
        "gemini-3.5-flash",
        "gemini-3.7-flash",
        "gemini-3.8-flash",
    ]
    seen = set()
    models_to_try = [m for m in candidate_models if not (m in seen or seen.add(m))]

    user_content = build_user_message(user_prompt, headers)
    last_error = None

    for model in models_to_try:
        try:
            response = client.models.generate_content(
                model=model,
                contents=user_content,
                config=types.GenerateContentConfig(
                    system_instruction=SYSTEM_PROMPT,
                    response_mime_type="application/json",
                    temperature=0.0,
                ),
            )
            text = response.text or "{}"
            parsed = _clean_and_parse_json(text)
            print(f"[Using Google Gemini LLM ({model})]")
            return parsed
        except Exception as err:
            last_error = err
            continue

    raise RuntimeError(f"All Gemini models failed. Last error: {last_error}")


def _call_openai(api_key: str, user_prompt: str, headers: list[str] | None = None) -> dict[str, Any]:
    """Calls OpenAI API."""
    from openai import OpenAI

    client = OpenAI(api_key=api_key)
    model = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
    user_content = build_user_message(user_prompt, headers)

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_content},
        ],
        temperature=0.0,
        response_format={"type": "json_object"},
    )

    content = response.choices[0].message.content or "{}"
    print(f"[Using OpenAI LLM ({model})]")
    return _clean_and_parse_json(content)


def understand_request(user_prompt: str, headers: list[str] | None = None) -> dict[str, Any]:
    """
    Sends the user's natural-language instruction to the LLM and returns parsed JSON operations.
    
    Priority:
      1. GEMINI_API_KEY or GOOGLE_API_KEY (Google Gemini)
      2. OPENAI_API_KEY (OpenAI)
      3. Fallback parser (if no active API keys are found)
    """
    load_dotenv(override=True)

    # 1. Check Gemini
    gemini_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if gemini_key and gemini_key.strip() and "your_" not in gemini_key:
        try:
            return _call_gemini(gemini_key.strip(), user_prompt, headers)
        except Exception as e:
            print(f"[Warning] Gemini API call failed ({e}). Falling back to local instruction parser.")
            return _fallback_parse(user_prompt)

    # 2. Check OpenAI
    openai_key = os.getenv("OPENAI_API_KEY")
    if openai_key and openai_key.strip() and "your_" not in openai_key:
        try:
            return _call_openai(openai_key.strip(), user_prompt, headers)
        except Exception as e:
            print(f"[Warning] OpenAI API call failed ({e}). Falling back to local instruction parser.")
            return _fallback_parse(user_prompt)

    # 3. Offline / local fallback
    print("[Notice] No active GEMINI_API_KEY or OPENAI_API_KEY detected. Using local instruction parser.")
    return _fallback_parse(user_prompt)
