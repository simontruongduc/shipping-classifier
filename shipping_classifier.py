# -*- coding: utf-8 -*-
import pandas as pd
import os
import sys
import re

REQUIRED_HEADERS = ["Container NO", "UN Type", "HTS CODE", "Desc.of Goods"]

FACTORY_RULES = {
    "OP": {
        "DG": {
            "values": ["DG"],
            "footer": [
                "DG GOODS",
                "UN No: UN3481",
                "Technical name: LITHIUM ION BATTERIES ARE PACKED WITH EQUIPMENT",
                "IMO/CRF class: 9",
                "UN PACKING CODE: 4G",
            ],
        },
        "NONDG": {
            "values": ["NONDG"],
            "footer": ["NONDG GOODS CONTAIN BATTERY"],
        },
        "GENERAL": {
            "values": ["GENERAL"],
            "footer": ["GENERAL GOODS WITHOUT BATTERY"],
        },
    },
    "CPT": {
        "DG": None,  # CPT không có DG
        "NONDG": {
            "values": ["NON-DG"],
            "footer": ["Non-DG GOODS"],
        },
        "GENERAL": {
            "values": ["GENERAL CARGO"],
            "footer": ["General Cargo"],
        },
    },
    "MILWAUKEE": {
        "DG": {
            "values": [
                "UN3481 LITHIUM BATTERY PACKED WITH EQUIPMENT CLASS 9 PG N/A – PI 966"
            ],
            "footer": [
                "UN3481 LITHIUM BATTERY PACKED WITH EQUIPMENT CLASS 9 PG N/A – PI 966"
            ],
        },
        "NONDG": {
            "values": ["NON-DG WITH BATTERY"],
            "footer": ["NON-DG WITH BATTERY"],
        },
        "NONDG-WITHOUT": {
            "values": ["NON-DG WITHOUT BATTERY"],
            "footer": ["NON-DG WITHOUT BATTERY"],
        },
        "GENERAL": {
            "values": ["GENERAL CARGO WITHOUT BATTERY"],
            "footer": ["GENERAL CARGO WITHOUT BATTERY"],
        },
    },
}


def normalize_name(s: str) -> str:
    """Chuẩn hóa tên cột / chuỗi: lower, remove BOM, keep alnum, collapse spaces"""
    if pd.isna(s):
        return ""
    s = str(s).replace("\ufeff", "").strip().lower()
    # thay tất cả ký tự non-alnum thành space
    s = re.sub(r"[^0-9a-z]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def find_header_row(file_path, ext, sheet_name=None, preview_rows=30):
    """
    Tìm dòng header bằng cách scan preview_rows đầu tiên.
    Trả về index của header (0-based). Nếu không tìm thấy, trả về 0.
    """
    if ext == ".csv":
        preview_df = pd.read_csv(file_path, encoding="utf-8-sig", header=None, nrows=preview_rows)
    else:
        preview_df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl", header=None, nrows=preview_rows)

    target_norms = {normalize_name(h): h for h in REQUIRED_HEADERS}

    for i in range(len(preview_df)):
        row = preview_df.iloc[i].astype(str).tolist()
        row_norms = [normalize_name(v) for v in row]
        # kiểm tra tất cả required header có trong dòng này không
        if all(t in row_norms for t in target_norms):
            return i
    return 0


def choose_input_file():
    while True:
        file_path = input("📂 Nhập đường dẫn file CSV hoặc Excel: ").strip('"').strip()
        if os.path.exists(file_path):
            return file_path
        print("❌ Lỗi: File không tồn tại. Vui lòng nhập lại!")


def choose_excel_sheet(file_path):
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheets = xls.sheet_names
    print("\n📌 Danh sách các sheet trong file:")
    for i, sheet in enumerate(sheets, start=1):
        print(f"   {i}. {sheet}")
    choice = input("\n🔹 Nhập số thứ tự sheet (Enter = sheet đầu tiên): ").strip()
    return sheets[int(choice) - 1] if choice.isdigit() and 1 <= int(choice) <= len(sheets) else sheets[0]


def choose_factory_type():
    factories = list(FACTORY_RULES.keys())
    print("\n🏭 Chọn loại nhà máy cần xử lý:")
    for i, fac in enumerate(factories, start=1):
        print(f"   {i}. {fac}")
    choice = input("\n🔹 Nhập số thứ tự (1-{}): ".format(len(factories))).strip()
    return factories[int(choice) - 1] if choice.isdigit() and 1 <= int(choice) <= len(factories) else factories[0]


def write_section(f, container_no, df, footer):
    if df.empty:
        return
    f.write(f"cont: {container_no}\n")
    for desc in df["Desc.of Goods"].dropna().tolist():
        f.write(f"{desc}\n")
    hs_codes = (
        df["HTS CODE"]
        .dropna()
        .apply(lambda x: str(int(x)) if isinstance(x, (int, float)) and x == int(x) else str(x))
        .unique()
        .tolist()
    )

    if hs_codes:
        f.write(f"hs code: {', '.join(hs_codes)}\n")
    for line in footer:
        f.write(f"{line}\n")
    f.write("---------------------------------\n")


def process_file(input_file, output_file, factory_type):
    try:
        ext = os.path.splitext(input_file)[1].lower()
        sheet_name = None
        if ext in [".xls", ".xlsx"]:
            sheet_name = choose_excel_sheet(input_file)

        header_row = find_header_row(input_file, ext, sheet_name)

        # đọc file với skiprows = header_row để header trở thành hàng đầu tiên
        if ext == ".csv":
            df = pd.read_csv(input_file, encoding="utf-8-sig", skiprows=header_row, header=0)
        else:
            df = pd.read_excel(input_file, sheet_name=sheet_name, engine="openpyxl", skiprows=header_row, header=0)

        # Chuẩn hóa và map tên cột
        col_map = {}
        target_norms = {normalize_name(h): h for h in REQUIRED_HEADERS}
        for col in df.columns:
            n = normalize_name(col)
            if n in target_norms:
                col_map[col] = target_norms[n]
        if col_map:
            df = df.rename(columns=col_map)

        # Kiểm tra cột bắt buộc
        for col in REQUIRED_HEADERS:
            if col not in df.columns:
                print(f"❌ Lỗi: Không tìm thấy cột '{col}' trong file.")
                sys.exit(1)

        df = df.dropna(how="all")

        # Điền Container NO bị trống (carry-forward)
        container_value = None
        for i in df.index:
            current_val = str(df.at[i, "Container NO"]).strip()
            if current_val and current_val.lower() != "nan":
                container_value = current_val
            elif container_value:
                df.at[i, "Container NO"] = container_value

        # Nhóm theo container
        grouped = {}
        for _, row in df.iterrows():
            container_no = str(row["Container NO"]).strip()
            grouped.setdefault(container_no, []).append(row)

        if os.path.exists(output_file):
            os.remove(output_file)

        with open(output_file, "w", encoding="utf-8") as f:
            for container_no, rows in grouped.items():
                container_df = pd.DataFrame(rows).drop_duplicates(subset=["UN Type", "Desc.of Goods"])
                rules = FACTORY_RULES[factory_type]

                for section, rule in rules.items():
                    if not rule:
                        continue
                    allowed = [v.upper() for v in rule["values"]]
                    mask = container_df["UN Type"].astype(str).str.upper().isin(allowed)
                    section_df = container_df[mask]
                    write_section(f, container_no, section_df, rule["footer"])

        print(f"\n✅ Đã xử lý xong! File xuất: {output_file}")

    except Exception as e:
        print("❌ Lỗi:", e)


if __name__ == "__main__":
    input_file = choose_input_file()
    base_dir = os.path.dirname(os.path.abspath(input_file))
    file_name = os.path.splitext(os.path.basename(input_file))[0]
    output_file = os.path.join(base_dir, f"{file_name}_output.txt")

    factory_type = choose_factory_type()
    process_file(input_file, output_file, factory_type)
