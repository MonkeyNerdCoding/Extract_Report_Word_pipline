import os
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import datetime
from docx.oxml import OxmlElement, ns
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  


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


def set_table_borders(table):
    tbl = table._element
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement(f"w:{border_name}")
        border.set(ns.qn("w:val"), "single")
        border.set(ns.qn("w:sz"), "4")     
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

    # Footer
    for section in doc.sections:
        footer = section.footer
        for p in footer.paragraphs:
            for run in p.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, replacement)


def create_pie_chart(df, title, output_image, label_col_idx=0, value_col_idx=1, top_n=10):
    """
    Tạo pie chart từ DataFrame với style giống hình mẫu.
    
    Args:
        df: DataFrame chứa data
        title: Tiêu đề chart
        output_image: Đường dẫn file image output
        label_col_idx: Index của cột label (default 0)
        value_col_idx: Index của cột value (default 1)
        top_n: Số lượng items hiển thị (default 10)
    """
    try:
        # Lấy top N rows
        df_chart = df.head(top_n).copy()
        
        # Lấy data cho chart
        labels = df_chart.iloc[:, label_col_idx].astype(str).tolist()
        values = pd.to_numeric(df_chart.iloc[:, value_col_idx], errors='coerce').fillna(0).tolist()
        
        # Debug: Print data info
        print(f"   🔍 Chart data preview for '{title}':")
        print(f"      Labels: {labels[:3]}...")
        print(f"      Values: {values[:3]}...")
        
        # Filter out invalid/zero values
        valid_data = [(label, value) for label, value in zip(labels, values) 
                      if value > 0 and label.strip() != '' and label.lower() not in ['nan', 'none']]
        
        if not valid_data:
            print(f"   ⚠️ No valid data for chart: {title} (all values are 0, NaN, or invalid)")
            return False
        
        # Unzip filtered data
        labels, values = zip(*valid_data)
        
        # Tạo figure với nền trắng
        fig, ax = plt.subplots(figsize=(10, 8), facecolor='white')
        
        # Định nghĩa màu sắc giống hình mẫu (màu đậm hơn, professional)
        colors = [
            '#5B9BD5',  # Blue
            '#ED7D31',  # Orange/Red
            '#A5A5A5',  # Gray
            '#FFC000',  # Yellow
            '#70AD47',  # Green
            '#4472C4',  # Dark Blue
            '#C55A11',  # Dark Orange
            '#7030A0',  # Purple
            '#44546A',  # Dark Gray
            '#264478',  # Navy
        ]
        
        wedges, texts = ax.pie(
            values,
            labels=None,  # Không hiển thị labels trên pie
            startangle=90,
            colors=colors[:len(values)],
            explode=[0.02] * len(values)  # Tách nhẹ các phần
        )
        
        ax.set_title(title, fontsize=16, fontweight='bold', pad=30, color='#333333')
        
        num_cols = min(3, len(labels))  
        if len(labels) > 6:
            num_cols = 3
        elif len(labels) > 3:
            num_cols = 2
        else:
            num_cols = 1
            
        legend = ax.legend(
            labels,
            loc='upper center',
            bbox_to_anchor=(0.5, -0.05),
            ncol=num_cols,
            frameon=False,
            fontsize=11,
            handlelength=1.5,
            handleheight=1.5,
            columnspacing=2
        )
        
        try:
            for i, wedge in enumerate(wedges):
                if i < len(legend.get_patches()):
                    legend.get_patches()[i].set_facecolor(colors[i % len(colors)])
        except:
            pass
        
        plt.axis('equal')
        plt.tight_layout()
        
        # Save image với độ phân giải cao, nền trắng
        plt.savefig(output_image, dpi=150, bbox_inches='tight', facecolor='white')
        plt.close()
        
        print(f"   📊 Created chart: {output_image}")
        return True
        
    except Exception as e:
        print(f"   ❌ Error creating chart: {e}")
        import traceback
        traceback.print_exc()
        plt.close()
        return False


def generate_report(excel_file: str, template_file: str, output_file: str, mapping: dict, chart_mapping: dict = None):
    xls = pd.ExcelFile(excel_file)
    doc = Document(template_file)
    temp_images = []  # Track temp image files để cleanup sau

    # Process regular table placeholders
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

                    p._element.getparent().replace(p._element, table._element)

            print(f"✅ Replaced {placeholder} with sheet '{sheet_name}' (rows={len(df)})")

        except Exception as e:
            print(f"⚠️ Could not process {placeholder}: {e}")

    # Process chart placeholders
    if chart_mapping:
        for placeholder, config in chart_mapping.items():
            try:
                sheet_name = config["sheet"]
                chart_title = config.get("title", sheet_name)
                
                # Read data
                df = pd.read_excel(xls, sheet_name=sheet_name)
                
                # Create temporary image
                temp_image = f"temp_chart_{placeholder.strip('<>').replace('_', '')}.png"
                temp_images.append(temp_image)
                
                # Create pie chart
                label_col = config.get("label_col", 0)
                value_col = config.get("value_col", 1)
                top_n = config.get("top_n", 10)
                
                if create_pie_chart(df, chart_title, temp_image, label_col, value_col, top_n):
                    # Insert into Word
                    for p in doc.paragraphs:
                        if placeholder in p.text:
                            # Clear placeholder text
                            p.text = p.text.replace(placeholder, "")
                            # Add image
                            run = p.add_run()
                            run.add_picture(temp_image, width=Inches(5.5))
                            print(f"✅ Inserted chart for {placeholder}")
                            break
                            
            except Exception as e:
                print(f"⚠️ Could not create chart for {placeholder}: {e}")

    # Try to save with error handling
    try:
        doc.save(output_file)
        print(f"\n📄 Report generated: {output_file}")
    except PermissionError:
        # Nếu file đang mở, thử tạo file mới với suffix
        import time
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = os.path.splitext(output_file)[0]
        new_output = f"{base_name}_{timestamp}.docx"
        doc.save(new_output)
        print(f"\n⚠️ File gốc đang được mở. Đã lưu thành: {new_output}")
    
    # Cleanup temporary images
    for temp_img in temp_images:
        try:
            if os.path.exists(temp_img):
                os.remove(temp_img)
        except:
            pass


if __name__ == "__main__":
    template_folder = r"D:\INTERNSHIP\SQL_merge_260112\SQL_merge\SQL_merge\rptemplate"
    excel_folder = r"D:\INTERNSHIP\SQL_merge_260112\SQL_merge\SQL_merge\output"
    output_folder = r"D:\INTERNSHIP\SQL_merge_260112\SQL_merge\SQL_merge\reports"
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

    # Chart mapping for pie charts
    chart_mapping = {
        "<cpu_usage_chart>": {
            "sheet": "CPU Usage by Database",
            "title": "Chart 1. CPU Usage by Database",
            "label_col": 1,  # Column B: Database Name
            "value_col": 3,  # Column D: CPU Percent
            "top_n": 10
        },
        "<io_usage_chart>": {
            "sheet": "IO Usage By Database",
            "title": "Chart 2. IO Usage By Database",
            "label_col": 1,  # Column B: Database Name
            "value_col": 3,  # Column D: Total I/O %
            "top_n": 10
        },
        "<buffer_usage_chart>": {
            "sheet": "Total Buffer Usage by Database",
            "title": "Chart 3. Total Buffer Usage by Database",
            "label_col": 1,  # Column B: Database Name
            "value_col": 3,  # Column D: Buffer Pool Percent
            "top_n": 10
        },
    }

    for template_file in os.listdir(template_folder):
        # Skip temporary Word files and non-docx files
        if not template_file.lower().endswith(".docx") or template_file.startswith("~$"):
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
            # Skip temporary Excel files
            if keyword in excel_file and excel_file.lower().endswith(".xlsx") and not excel_file.startswith("~$"):
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
            chart_mapping=chart_mapping,
        )