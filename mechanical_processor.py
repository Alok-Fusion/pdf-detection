import os
import re
from collections import defaultdict

import pdfplumber
import fitz  # PyMuPDF
import pandas as pd


# ------------ CONFIG ------------

EXCEL_PATH = "mechanical_schedule_summary.xlsx"

# Light pastel color palette (RGB 0–1)
LIGHT_PALETTE = [
    (0.75, 0.45, 0.45),  # muted red
    (0.45, 0.45, 0.75),  # muted blue
    (0.45, 0.75, 0.45),  # muted green
    (0.75, 0.75, 0.45),  # muted yellow
    (0.75, 0.55, 0.35),  # muted orange
    (0.65, 0.45, 0.75),  # muted purple
    (0.35, 0.65, 0.65),  # muted cyan
    (0.75, 0.65, 0.45),  # muted peach
    (0.45, 0.65, 0.75),  # muted sky
    (0.45, 0.75, 0.55), 
]


# ------------ HELPERS ------------

def mark_type(mark: str) -> str:
    """
    Get the 'type' of a mark.
    Example: FCU-10 -> FCU, L-1 -> L
    """
    if not mark:
        return ""
    m = re.match(r'^[A-Z]+', mark)
    if m:
        return m.group(0)
    if "-" in mark:
        return mark.split("-")[0]
    return mark


def get_plan_type(page) -> str | None:
    """
    Detect which plan this page belongs to, based on text.
    """
    txt = page.get_text()
    lines = txt.splitlines()
    plan_lines = [line.strip() for line in lines if "PLAN" in line.upper()]

    # Prefer mechanical plan labels
    mech_lines = [l for l in plan_lines if "MECHANICAL" in l.upper()]
    if mech_lines:
        return mech_lines[0]

    for key in ["FLOOR PLAN", "PIPING PLAN", "DETECTION PLAN"]:
        for l in plan_lines:
            if key in l.upper():
                return l

    if plan_lines:
        return plan_lines[0]

    return None


def extract_schedules_and_marks(pdf_path: str):
    """
    Extract schedule tables and marks from a PDF.
    Returns:
      schedule_tables: list of dicts
      marks: sorted list of unique marks
    """
    schedule_tables = []
    marks_set = set()

    with pdfplumber.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf.pages):
            text_upper = (page.extract_text() or "").upper()
            page_has_schedule_word = "SCHEDULE" in text_upper
            tables = page.extract_tables()

            for t_index, table in enumerate(tables):
                if not table or all(
                    all((cell is None or str(cell).strip() == "") for cell in row)
                    for row in table
                ):
                    continue

                # Find header row containing "MARK"
                header_row_idx = None
                for r_idx, row in enumerate(table):
                    if any(
                        (str(c).strip().upper().find("MARK") != -1)
                        for c in row if c is not None
                    ):
                        header_row_idx = r_idx
                        break

                # If no MARK and no SCHEDULE on page, skip
                if header_row_idx is None and not page_has_schedule_word:
                    continue

                # Schedule name: look at rows above header
                schedule_name = ""
                if header_row_idx is not None:
                    for up in range(header_row_idx - 1, -1, -1):
                        row_up = table[up]
                        cells = [str(c).strip() for c in row_up if c not in [None, ""]]
                        if cells:
                            schedule_name = cells[0]
                            break

                # Build header & data rows
                header = None
                data_rows = []

                if header_row_idx is not None:
                    raw_header = table[header_row_idx]
                    header = [str(c).strip() if c is not None else "" for c in raw_header]

                    for r in range(header_row_idx + 1, len(table)):
                        row = table[r]
                        if any(
                            cell not in [None, ""]
                            and str(cell).strip() != ""
                            for cell in row
                        ):
                            data_rows.append(
                                [str(c).strip() if c is not None else "" for c in row]
                            )
                else:
                    for row in table:
                        if header is None and any(
                            cell not in [None, ""]
                            and str(cell).strip() != ""
                            for cell in row
                        ):
                            header = [
                                str(c).strip() if c is not None else "" for c in row
                            ]
                        else:
                            if any(
                                cell not in [None, ""]
                                and str(cell).strip() != ""
                                for cell in row
                            ):
                                data_rows.append(
                                    [str(c).strip() if c is not None else "" for c in row]
                                )

                if not header or not data_rows:
                    continue

                header_upper = [h.upper() for h in header]
                looks_like_schedule = page_has_schedule_word or (
                    "SCHEDULE" in " ".join(header_upper)
                )

                if not looks_like_schedule:
                    continue

                # Convert rows to dicts
                table_dict_rows = []
                for row in data_rows:
                    row_dict = {}
                    for col_idx, col_name in enumerate(header):
                        key = col_name if col_name else f"COL_{col_idx+1}"
                        value = row[col_idx] if col_idx < len(row) else ""
                        row_dict[key] = value
                    table_dict_rows.append(row_dict)

                schedule_tables.append(
                    {
                        "file_name": os.path.basename(pdf_path),
                        "page_index": page_index,
                        "table_index_on_page": t_index,
                        "schedule_name": schedule_name,
                        "header": header,
                        "rows": table_dict_rows,
                    }
                )

                # Collect mark values from any column containing "MARK"
                mark_col_indices = [
                    i for i, h in enumerate(header_upper) if "MARK" in h
                ]
                for mark_col in mark_col_indices:
                    for row in data_rows:
                        if mark_col < len(row):
                            mark_val = row[mark_col].strip()
                            if mark_val:
                                marks_set.add(mark_val)

    return schedule_tables, sorted(list(marks_set))


def highlight_pdf_and_count(pdf_path: str, marks: list[str], output_dir: str):
    """
    Highlight all marks in the given PDF with light colors per mark type.
    Returns:
      highlighted_path, plan_counts_rows (list of dicts)
    """
    base_name = os.path.basename(pdf_path)
    os.makedirs(output_dir, exist_ok=True)
    highlighted_path = os.path.join(
        output_dir, os.path.splitext(base_name)[0] + "_highlighted_pastel.pdf"
    )

    # Map type -> color
    types = sorted(set(mark_type(m) for m in marks if m))
    type_color_map = {
        t: LIGHT_PALETTE[idx % len(LIGHT_PALETTE)]
        for idx, t in enumerate(types)
    }

    doc = fitz.open(pdf_path)
    plan_counts = defaultdict(lambda: defaultdict(int))  # plan -> mark -> count

    for page in doc:
        plan = get_plan_type(page)
        for mark in marks:
            if not mark:
                continue

            # Basic search for exact text
            rects = page.search_for(mark)

            # Optionally: try normalizing hyphens / spaces if you see misses
            if not rects and "-" in mark:
                alt = mark.replace("-", "-")  # common different hyphen
                rects = page.search_for(alt)

            if not rects:
                continue

            t = mark_type(mark)
            color = type_color_map.get(t, (1, 1, 0.9))  # default light yellow
            for rect in rects:
                annot = page.add_highlight_annot(rect)
                annot.set_colors(stroke=color)
                annot.update()

            if plan:
                plan_counts[plan][mark] += len(rects)

    doc.save(highlighted_path)
    doc.close()

    # Convert counts to flat rows
    plan_count_rows = []
    for plan, mark_dict in plan_counts.items():
        for mark, count in mark_dict.items():
            plan_count_rows.append(
                {
                    "file_name": base_name,
                    "plan": plan,
                    "mark": mark,
                    "mark_type": mark_type(mark),
                    "count": count,
                }
            )

    return highlighted_path, plan_count_rows


def update_excel(schedule_tables: list[dict], plan_count_rows: list[dict], excel_path: str = EXCEL_PATH):
    """
    Update (or create) the Excel file with schedule details + plan counts.
    - Sheet 'Schedules': all schedule table rows (with columns unioned)
    - Sheet 'PlanCounts': counts of each mark on each plan
    """
    # Build rows for Schedules
    schedule_rows = []
    for table in schedule_tables:
        file_name = table["file_name"]
        page_index = table["page_index"]
        table_index = table["table_index_on_page"]
        schedule_name = table["schedule_name"]
        for row_dict in table["rows"]:
            out_row = {
                "file_name": file_name,
                "page_index": page_index,
                "table_index": table_index,
                "schedule_name": schedule_name,
            }
            out_row.update(row_dict)
            schedule_rows.append(out_row)

    new_schedules_df = pd.DataFrame(schedule_rows)
    new_plan_counts_df = pd.DataFrame(plan_count_rows)

    if os.path.exists(excel_path):
        try:
            existing_schedules = pd.read_excel(excel_path, sheet_name="Schedules")
        except Exception:
            existing_schedules = pd.DataFrame()

        try:
            existing_plan_counts = pd.read_excel(excel_path, sheet_name="PlanCounts")
        except Exception:
            existing_plan_counts = pd.DataFrame()
    else:
        existing_schedules = pd.DataFrame()
        existing_plan_counts = pd.DataFrame()

    # Append
    if not existing_schedules.empty:
        schedules_df = pd.concat(
            [existing_schedules, new_schedules_df], ignore_index=True
        )
    else:
        schedules_df = new_schedules_df

    if not existing_plan_counts.empty:
        plan_counts_df = pd.concat(
            [existing_plan_counts, new_plan_counts_df], ignore_index=True
        )
    else:
        plan_counts_df = new_plan_counts_df

    # Save
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        schedules_df.to_excel(writer, sheet_name="Schedules", index=False)
        plan_counts_df.to_excel(writer, sheet_name="PlanCounts", index=False)

    return excel_path
