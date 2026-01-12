import os
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import datetime
from docx.oxml import OxmlElement, ns


def set_cell_bg(cell, fill_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill_color)
    tcPr.append(shd)


def format_cell(cell, bold=False, font_color=None):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Cambria"
            run.font.size = Pt(12)
            run.bold = bold
            if font_color:
                run.font.color.rgb = font_color

from docx.oxml import OxmlElement

def set_table_borders(table):
    tbl = table._element
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement(f"w:{border_name}")
        border.set(ns.qn("w:val"), "single")
        border.set(ns.qn("w:sz"), "4")      # độ dày 0.5pt
        border.set(ns.qn("w:space"), "0")
        border.set(ns.qn("w:color"), "000000")
        tblBorders.append(border)
    tbl.tblPr.append(tblBorders)

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._element = None

def replace_placeholder_text(doc, placeholder, replacement):
    """Thay thế placeholder trong paragraphs và table cells."""
    # Paragraphs
    for p in doc.paragraphs:
        for run in p.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, replacement)

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, replacement)

    #footer
    for section in doc.sections:
        footer = section.footer
        for p in footer.paragraphs:
            for run in p.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, replacement)


def generate_report(excel_file: str, template_file: str, output_file: str, mapping: dict):
    xls = pd.ExcelFile(excel_file)
    doc = Document(template_file)

    for placeholder, config in mapping.items():
        try:
            # ---- Special case cho collect_date ----
            if placeholder == "<collect_date>":
                current_date = datetime.now().strftime("%m.%Y")  # hoặc "Tháng %m.%Y"
                replace_placeholder_text(doc, placeholder, current_date)
                print(f"✅ Replaced {placeholder} with {current_date}")
                continue
            # ---------------------------------------

            if not config:
                continue

            sheet_name = config["sheet"]
            df = pd.read_excel(xls, sheet_name=sheet_name)

            if "columns" in config and config["columns"]:
                col_indices = config["columns"]
                selected = [df.columns[i] for i in col_indices if i < len(df.columns)]
                df = df[selected]

            max_rows = config.get("max_rows", None)
            if max_rows and len(df) > max_rows:
                df = df.head(max_rows)

            for p in doc.paragraphs:
                if placeholder in p.text:
                    table = doc.add_table(rows=1, cols=len(df.columns))
                    table.autofit = True
                    set_table_borders(table)   # 👈 thêm viền

                    hdr_cells = table.rows[0].cells
                    for j, col in enumerate(df.columns):
                        hdr_cells[j].text = str(col)
                        set_cell_bg(hdr_cells[j], "0066CC")
                        format_cell(hdr_cells[j], bold=True, font_color=RGBColor(255, 255, 255))

                    for _, row in df.iterrows():
                        row_cells = table.add_row().cells
                        for j, val in enumerate(row):
                            row_cells[j].text = str(val)
                            format_cell(row_cells[j])

                    #p.text = p.text.replace(placeholder, "")
                    #p._element.addnext(table._element)
                    p._element.getparent().replace(p._element, table._element)

            print(f"✅ Replaced {placeholder} with sheet '{sheet_name}' (rows={len(df)})")

        except Exception as e:
            print(f"⚠️ Could not process {placeholder}: {e}")

    doc.save(output_file)
    print(f"\n📄 Report generated: {output_file}")


if __name__ == "__main__":
    template_folder = r"D:\SQL_merge\SQL_merge\rptemplate"
    excel_folder = r"D:\SQL_merge\SQL_merge\output"
    output_folder = r"D:\SQL_merge\SQL_merge\reports"
    os.makedirs(output_folder, exist_ok=True)

    mapping = {
        "<file_size>": {"sheet": "File Sizes and Space", "columns": [0, 1, 2, 3, 4, 5, 7], "max_rows": 50},
        "<fileio>": {"sheet": "IO Stats By File", "max_rows": 50},
        "<conn_count>": {"sheet": "Connection Counts by IP Address", "max_rows": 50},
        "<cpu_usage>": {"sheet": "CPU Usage by Database", "columns": [0, 1, 3], "max_rows": 50},
        "<io_usage>": {"sheet": "IO Usage By Database", "columns": [0, 1, 3], "max_rows": 50},
        "<buffer_usage>": {"sheet": "Total Buffer Usage by Database", "columns": [0, 1, 3], "max_rows": 50},
        "<top_worker>": {"sheet": "Top Worker Time Queries", "columns": [0, 1, 2, 4], "max_rows": 50},
        "<missing_index>": {"sheet": "Missing Indexes", "columns": [2, 5, 6, 7, 9], "max_rows": 50},
        "<agent_job>": {"sheet": "SQL Server Agent Jobs", "columns": [0, 1, 2, 3, 4, 8, 9], "max_rows": 50},
        "<recent_bk>": {"sheet": "Recent Full Backups", "columns": [2, 3, 4, 5, 11], "max_rows": 50},
        "<collect_date>": {},
    }

    for template_file in os.listdir(template_folder):
        if not template_file.lower().endswith(".docx"):
            continue

        base_name = os.path.splitext(template_file)[0]
        keyword = None
        for part in base_name.split("_"):
            if part.startswith("INS"):
                keyword = part
                break

        if not keyword:
            print(f"⚠️ Không tìm thấy keyword 'INS...' trong {template_file}")
            continue

        excel_match = None
        for excel_file in os.listdir(excel_folder):
            if keyword in excel_file and excel_file.lower().endswith(".xlsx"):
                excel_match = excel_file
                break

        if not excel_match:
            print(f"⚠️ Không tìm thấy excel cho template {template_file} (keyword={keyword})")
            continue

        output_file = os.path.join(output_folder, f"{template_file}")

        generate_report(
            excel_file=os.path.join(excel_folder, excel_match),
            template_file=os.path.join(template_folder, template_file),
            output_file=output_file,
            mapping=mapping,
        )
