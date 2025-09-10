# -*- coding: utf-8 -*-
import pandas as pd
import os
import sys

# Danh sách header chuẩn
EXPECTED_HEADERS = [
    "SO#", "Container type", "Container NO", "Seal NO", "QTY", "CTN", "No of Package",
    "Total No of Package", "N.W.", "GW", "CBM", "UN Type", "HTS CODE", "END BYR PO",
    "TTI MODEL", "CUST MODEL", "OFIS LN NBR", "OFIS SHIPMENT NBR", "OFIS LINE ID",
    "CSR Number", "Desc.of Goods"
]

def choose_input_file():
    """Hỏi người dùng nhập đường dẫn file"""
    while True:
        file_path = input("📂 Nhập đường dẫn file CSV hoặc Excel: ").strip('"').strip()
        if os.path.exists(file_path):
            return file_path
        else:
            print("❌ Lỗi: File không tồn tại. Vui lòng nhập lại!")

def choose_excel_sheet(file_path):
    """Hỏi người dùng chọn sheet khi mở Excel"""
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheets = xls.sheet_names

    print("\n📌 Danh sách các sheet trong file:")
    for i, sheet in enumerate(sheets, start=1):
        print(f"   {i}. {sheet}")

    choice = input("\n🔹 Nhập số thứ tự sheet muốn đọc (Enter để chọn sheet đầu tiên): ").strip()

    if choice == "":
        return sheets[0]

    if choice.isdigit() and 1 <= int(choice) <= len(sheets):
        return sheets[int(choice) - 1]

    print("⚠️ Lựa chọn không hợp lệ. Mặc định đọc sheet đầu tiên.")
    return sheets[0]

def find_header_row(file_path, ext, sheet_name=None):
    """
    Xác định dòng header trong file Excel hoặc CSV
    Trả về index của header
    """
    if ext == ".csv":
        preview_df = pd.read_csv(file_path, encoding="utf-8-sig", header=None, nrows=30)
    else:
        preview_df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl", header=None, nrows=30)

    for i in range(len(preview_df)):
        row_values = preview_df.iloc[i].astype(str).str.strip().tolist()
        # Nếu số cột trùng khớp và các header chính xuất hiện trong dòng này
        matches = [h for h in EXPECTED_HEADERS if h in row_values]
        if len(matches) >= 5:  # Chỉ cần tìm thấy ≥5 header trùng là đủ tin tưởng
            return i

    print("⚠️ Không tìm thấy header phù hợp. Sử dụng dòng đầu tiên làm header.")
    return 0

def process_file(input_file, output_file):
    try:
        ext = os.path.splitext(input_file)[1].lower()

        # Chọn sheet nếu là Excel
        sheet_name = None
        if ext in [".xls", ".xlsx"]:
            sheet_name = choose_excel_sheet(input_file)

        # Tìm dòng header chính xác
        header_row = find_header_row(input_file, ext, sheet_name)

        # Đọc file, bỏ qua các dòng trước header
        if ext == ".csv":
            df = pd.read_csv(input_file, encoding="utf-8-sig", skiprows=header_row)
        else:
            df = pd.read_excel(input_file, sheet_name=sheet_name, engine="openpyxl", skiprows=header_row)

        # Đặt tên cột theo EXPECTED_HEADERS
        df.columns = EXPECTED_HEADERS

        print(f"\n✅ Đã tìm thấy header tại dòng: {header_row + 1}")

        # Loại bỏ dòng trống
        df = df.dropna(how="all")

        # Kiểm tra cột bắt buộc
        required_cols = ["Container NO", "UN Type", "HTS CODE", "Desc.of Goods"]
        for col in required_cols:
            if col not in df.columns:
                print(f"❌ Lỗi: Không tìm thấy cột '{col}' trong file.")
                sys.exit(1)

        # Điền giá trị Container NO bị trống
        container_value = None
        for i in df.index:
            current_val = str(df.at[i, "Container NO"]).strip()
            if current_val and current_val.lower() != "nan":
                container_value = current_val
            elif container_value:
                df.at[i, "Container NO"] = container_value

        # Group dữ liệu theo Container NO
        grouped = {}
        for _, row in df.iterrows():
            container_no = str(row["Container NO"]).strip()
            grouped.setdefault(container_no, []).append(row)

        # Xóa file output cũ nếu tồn tại
        if os.path.exists(output_file):
            os.remove(output_file)

        # Mở file để ghi dữ liệu
        with open(output_file, "w", encoding="utf-8") as f:
            for container_no, rows in grouped.items():
                container_df = pd.DataFrame(rows)
                container_df = container_df.drop_duplicates(subset=["UN Type", "Desc.of Goods"])

                dg_df = container_df[container_df["UN Type"].str.upper() == "DG"]
                nondg_df = container_df[container_df["UN Type"].str.upper() == "NONDG"]
                general_df = container_df[container_df["UN Type"].str.upper() == "GENERAL"]

                # --- Xử lý DG ---
                if not dg_df.empty:
                    f.write("cont: {}\n".format(container_no))
                    for desc in dg_df["Desc.of Goods"].dropna().tolist():
                        f.write("{}\n".format(desc))
                    hs_codes = dg_df["HTS CODE"].dropna().astype(str).unique().tolist()
                    f.write("hs code: {}\n".format(", ".join(hs_codes)))
                    f.write("DG GOODS\n")
                    f.write("UN No: UN3481\n")
                    f.write("Technical name: LITHIUM ION BATTERIES ARE PACKED WITH EQUIPMENT\n")
                    f.write("IMO/CRF class: 9\n")
                    f.write("UN PACKING CODE: 4G\n")
                    f.write("---------------------------------\n")

                # --- Xử lý NONDG ---
                if not nondg_df.empty:
                    f.write("cont: {}\n".format(container_no))
                    for desc in nondg_df["Desc.of Goods"].dropna().tolist():
                        f.write("{}\n".format(desc))
                    hs_codes = nondg_df["HTS CODE"].dropna().astype(str).unique().tolist()
                    f.write("hs code: {}\n".format(", ".join(hs_codes)))
                    f.write("NONDG GOODS CONTAIN BATTERY\n")
                    f.write("---------------------------------\n")

                # --- Xử lý GENERAL ---
                if not general_df.empty:
                    f.write("cont: {}\n".format(container_no))
                    for desc in general_df["Desc.of Goods"].dropna().tolist():
                        f.write("{}\n".format(desc))
                    hs_codes = general_df["HTS CODE"].dropna().astype(str).unique().tolist()
                    f.write("hs code: {}\n".format(", ".join(hs_codes)))
                    f.write("GENERAL GOODS WITHOUT BATTERY\n")
                    f.write("---------------------------------\n")

        print(f"\n✅ Đã xử lý xong! File xuất: {output_file}")

    except Exception as e:
        print("❌ Lỗi:", e)


if __name__ == "__main__":
    input_file = choose_input_file()
    base_dir = os.path.dirname(os.path.abspath(input_file))
    file_name = os.path.splitext(os.path.basename(input_file))[0]
    output_file = os.path.join(base_dir, f"{file_name}_output.txt")
    process_file(input_file, output_file)

