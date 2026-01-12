import os
import pandas as pd
import glob
from name_detect import extract_sheet_name

def merge_sql_csv(input_folder, output_file):
    csv_files = glob.glob(os.path.join(input_folder, "*.csv"))
    if not csv_files:
        print(f"⚠️ Không có CSV trong {input_folder}, bỏ qua.\n")
        return

    all_data = {}

    for file in csv_files:
        filename = os.path.basename(file)
        sheet_name = extract_sheet_name(filename)

        if not sheet_name:
            continue

        try:
            df = pd.read_csv(file)
            print(f"   ✅ {filename} ({len(df)} dòng)")
        except Exception as e:
            print(f"   ❌ Lỗi đọc {filename}: {e}")
            continue

        if sheet_name not in all_data:
            all_data[sheet_name] = []
        all_data[sheet_name].append(df)

    # Xuất Excel
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet_name, dfs in all_data.items():
            merged_df = pd.concat(dfs, ignore_index=True)
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"📝 Đã ghi sheet: {sheet_name} ({len(merged_df)} dòng)")

    print(f"✅ Done! File Excel sinh ra: {output_file}\n")


if __name__ == "__main__":
    parent_folder = r"D:\SQL_merge\SQL_merge\input"   # folder cha chứa nhiều DB
    output_folder = r"D:\SQL_merge\SQL_merge\output"

    os.makedirs(output_folder, exist_ok=True)

    for sub in os.listdir(parent_folder):
        sub_path = os.path.join(parent_folder, sub)
        if os.path.isdir(sub_path):  # chỉ xử lý folder con
            output_file = os.path.join(output_folder, f"{sub}_healthcheck_info.xlsx")
            print(f"\n🚀 Đang xử lý DB folder: {sub}")
            merge_sql_csv(sub_path, output_file)
