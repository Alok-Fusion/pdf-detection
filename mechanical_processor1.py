import os
import re
import json
from collections import defaultdict

import pdfplumber
import fitz  # PyMuPDF
import pandas as pd


# ================= CONFIG =================

OUTPUT_ROOT = "output"

# Light pastel color palette (RGB 0–1)
LIGHT_PALETTE = [
    (0.75, 0.45, 0.45),
    (0.45, 0.45, 0.75),
    (0.45, 0.75, 0.45),
    (0.75, 0.75, 0.45),
    (0.75, 0.55, 0.35),
    (0.65, 0.45, 0.75),
    (0.35, 0.65, 0.65),
    (0.75, 0.65, 0.45),
    (0.45, 0.65, 0.75),
]

MARK_REGEX = re.compile(r"[A-Z]{1,4}-\d+", re.IGNORECASE)


# ================= HELPERS =================

def mark_type(mark: str) -> str:
    if not mark:
        return ""
    m = re.match(r"^[A-Z]+", mark)
    return m.group(0) if m else mark


def build_search_variants(mark: str):
    return {
        mark,
        mark.replace("-", " "),
        mark.replace("-", ""),
        mark.lower(),
    }


def get_plan_type(page) -> str | None:
    text = page.get_text() or ""
    lines = [l.strip() for l in text.splitlines() if "PLAN" in l.upper()]

    for l in lines:
        if "MECHANICAL" in l.upper():
            return l
    return lines[0] if lines else None


def make_output_dirs(pdf_name: str):
    base = os.path.join(OUTPUT_ROOT, pdf_name)
    paths = {
        "base": base,
        "highlighted": os.path.join(base, "highlighted"),
        "data": os.path.join(base, "data"),
    }
    for p in paths.values():
        os.makedirs(p, exist_ok=True)
    return paths


# ================= EXTRACTION =================

def extract_schedules_and_marks(pdf_path: str):
    schedule_tables = []
    marks_set = set()

    with pdfplumber.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf.pages):
            text = page.extract_text() or ""

            # Extract marks via regex
            for m in MARK_REGEX.findall(text):
                marks_set.add(m.upper())

            # Extract schedule tables (basic)
            tables = page.extract_tables() or []
            for t_index, table in enumerate(tables):
                if not table:
                    continue

                header_row_idx = None
                for i, row in enumerate(table):
                    if any("MARK" in str(c).upper() for c in row if c):
                        header_row_idx = i
                        break

                if header_row_idx is None:
                    continue

                header = [str(c).strip() if c else "" for c in table[header_row_idx]]
                data_rows = table[header_row_idx + 1 :]

                rows = []
                for r in data_rows:
                    if any(c for c in r):
                        row_dict = {
                            header[i] if header[i] else f"COL_{i+1}":
                            str(r[i]).strip() if i < len(r) and r[i] else ""
                            for i in range(len(header))
                        }
                        rows.append(row_dict)

                schedule_tables.append(
                    {
                        "file_name": os.path.basename(pdf_path),
                        "page_index": page_index,
                        "table_index": t_index,
                        "header": header,
                        "rows": rows,
                    }
                )

    return schedule_tables, sorted(marks_set)


def extract_tags_from_excel(excel_path: str, tag_column: str):
    if excel_path.endswith(".csv"):
        df = pd.read_csv(excel_path)
    else:
        df = pd.read_excel(excel_path)

    if tag_column not in df.columns:
        raise ValueError(f"Column '{tag_column}' not found in Excel")

    tags = (
        df[tag_column]
        .dropna()
        .astype(str)
        .str.strip()
        .str.upper()
        .unique()
        .tolist()
    )
    return tags, df


# ================= HIGHLIGHTING =================

def highlight_pdf_and_count(pdf_path: str, tags: list[str]):
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
    dirs = make_output_dirs(pdf_name)

    highlighted_pdf_path = os.path.join(
        dirs["highlighted"], f"{pdf_name}_highlighted.pdf"
    )

    types = sorted(set(mark_type(t) for t in tags))
    type_color_map = {
        t: LIGHT_PALETTE[i % len(LIGHT_PALETTE)]
        for i, t in enumerate(types)
    }

    doc = fitz.open(pdf_path)

    plan_tag_counts = defaultdict(lambda: defaultdict(int))
    plan_type_counts = defaultdict(lambda: defaultdict(int))

    for page in doc:
        plan = get_plan_type(page)

        for tag in tags:
            rects = []
            for v in build_search_variants(tag):
                rects.extend(page.search_for(v))

            rects = list({r for r in rects})
            if not rects:
                continue

            color = type_color_map[mark_type(tag)]
            for r in rects:
                annot = page.add_highlight_annot(r)
                annot.set_colors(stroke=color)
                annot.update()

            if plan:
                plan_tag_counts[plan][tag] += len(rects)
                plan_type_counts[plan][mark_type(tag)] += len(rects)

    doc.save(highlighted_pdf_path)
    doc.close()

    return highlighted_pdf_path, plan_tag_counts, plan_type_counts, type_color_map, dirs


# ================= OUTPUT FILES =================

def write_summary_excel(path, tags, plan_tag, plan_type, excel_df=None):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame({"TAG": tags}).to_excel(writer, "All_Tags", index=False)
        pd.DataFrame(plan_tag).T.fillna(0).to_excel(writer, "Plan_by_Tag")
        pd.DataFrame(plan_type).T.fillna(0).to_excel(writer, "Plan_by_Type")

        if excel_df is not None:
            excel_df.to_excel(writer, "Source_Excel", index=False)


def write_json(path, data):
    with open(path, "w") as f:
        json.dump(data, f, indent=2)


# ================= MAIN PIPELINE =================

def run_pipeline(
    pdf_path: str,
    excel_path: str | None = None,
    excel_tag_column: str | None = None,
):
    if not pdf_path or not os.path.exists(pdf_path):
        raise ValueError("Valid PDF path is required")

    # ---- PDF extraction ----
    schedules, pdf_marks = extract_schedules_and_marks(pdf_path)

    # ---- Excel is OPTIONAL ----
    excel_tags = []
    excel_df = None

    if excel_path:
        if not excel_tag_column:
            raise ValueError(
                "excel_tag_column must be provided when excel_path is used"
            )
        if not os.path.exists(excel_path):
            raise ValueError("Excel file path does not exist")

        excel_tags, excel_df = extract_tags_from_excel(
            excel_path, excel_tag_column
        )

    # ---- Merge tags ----
    all_tags = sorted(set(pdf_marks + excel_tags))
    if not all_tags:
        raise ValueError("No tags found in PDF or Excel")

    # ---- Highlight & count ----
    highlighted_pdf, plan_tag, plan_type, colors, dirs = highlight_pdf_and_count(
        pdf_path, all_tags
    )

    # ---- Write Excel summary ----
    summary_excel_path = os.path.join(dirs["data"], "summary.xlsx")
    write_summary_excel(
        summary_excel_path,
        all_tags,
        plan_tag,
        plan_type,
        excel_df,
    )

    # ---- Write JSON ----
    json_path = os.path.join(dirs["data"], "data.json")
    write_json(
        json_path,
        {
            "pdf": os.path.basename(pdf_path),
            "total_tags": len(all_tags),
            "tags": all_tags,
            "plan_by_tag": plan_tag,
            "plan_by_type": plan_type,
            "colors": colors,
            "excel_used": bool(excel_path),
        },
    )

    return {
        "highlighted_pdf": highlighted_pdf,
        "summary_excel": summary_excel_path,
        "json": json_path,
        "output_dir": dirs["base"],
    }
