import os
import sqlite3
import pandas as pd
import warnings
from openpyxl import load_workbook

# Ẩn các cảnh báo của openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

DB_NAME = "master_data.db"
TABLE_NAME = "master_data"


def show_menu():
    print("\n=== MENU CHÍNH ===")
    print("1. Khởi tạo master data")
    print("2. Fill data")
    print("0. Thoát")
    return input("Nhập lựa chọn của bạn: ").strip()


def init_master_data():
    try:
        file_path = input("Nhập hoặc kéo thả file master data Excel vào đây: ").strip().strip('"').strip("'")

        if not os.path.exists(file_path):
            print("❌ File không tồn tại.")
            return

        print("Đang đọc file Excel...")
        excel = pd.ExcelFile(file_path)
        all_data = []

        for sheet_name in excel.sheet_names:
            df = excel.parse(sheet_name)

            required_cols = {
                'Carrier Bkg no.': 'D',
                'Job no.': 'E',
                'Shipper Name': 'F',
                'Carrier': 'N',
                'Vessel name': 'AD',
                'Origin ETD': 'AF',
                'New ETD': 'AG',
                'Status': 'H'
            }

            missing = [c for c in required_cols.keys() if c not in df.columns]
            if missing:
                print(f"⚠️ Sheet '{sheet_name}' bị thiếu cột: {missing}, bỏ qua.")
                continue

            # Ép kiểu & lọc dữ liệu
            df['Status'] = df['Status'].astype(str).str.lower()
            df = df[~df['Status'].isin(['cancel', 'shipper cancel'])]
            df = df[df['Carrier Bkg no.'].notna() & (df['Carrier Bkg no.'].astype(str).str.strip() != '')]

            if not df.empty:
                selected = df[list(required_cols.keys())]
                all_data.append(selected)

        if not all_data:
            print("❌ Không có dữ liệu hợp lệ trong file.")
            return

        master_df = pd.concat(all_data, ignore_index=True)
        master_df = master_df.fillna('').astype(str)

        # Ghi dữ liệu vào SQLite
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute(f"DROP TABLE IF EXISTS {TABLE_NAME}")
        conn.commit()

        master_df.to_sql(TABLE_NAME, conn, index=False, if_exists='replace', dtype={col: 'TEXT' for col in master_df.columns})
        cursor.execute(f"CREATE INDEX IF NOT EXISTS idx_carrier_bkg_no ON {TABLE_NAME}(`Carrier Bkg no.`)")
        conn.commit()
        conn.close()

        print("✅ Khởi tạo master data hoàn tất.")

    except KeyboardInterrupt:
        print("\n⚠️ Đã hủy quá trình khởi tạo (Ctrl + C). Quay lại menu...")
    except Exception as e:
        print(f"❌ Lỗi trong quá trình init master data: {e}")


def fill_data():
    try:
        if not os.path.exists(DB_NAME):
            print("⚠️ Chưa có master data. Hãy khởi tạo trước (chọn option 1).")
            return

        file_path = input("Nhập hoặc kéo thả file data cần xử lý: ").strip().strip('"').strip("'")

        if not os.path.exists(file_path):
            print("❌ File không tồn tại.")
            return

        import openpyxl

        # Đọc file input
        wb = openpyxl.load_workbook(file_path)
        sheets = wb.sheetnames

        print("\n📄 Danh sách sheet:")
        for i, s in enumerate(sheets, start=1):
            print(f"{i}. {s}")

        choice = input("Chọn sheet cần xử lý: ").strip()
        if not choice.isdigit() or int(choice) < 1 or int(choice) > len(sheets):
            print("❌ Lựa chọn không hợp lệ.")
            return

        sheet_name = sheets[int(choice) - 1]
        ws = wb[sheet_name]

        # Mở kết nối SQLite
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        print(f"🚀 Đang xử lý sheet: {sheet_name} ...")

        # Lấy header để xác định vị trí cột
        headers = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}

        # Kiểm tra cột BK# có tồn tại không
        if "BK #" not in headers:
            print("❌ Không tìm thấy cột 'BK #' trong sheet.")
            return

        bk_col = headers["BK #"]
        line_col = headers.get("LINE")
        job_col = headers.get("JOB")
        vessel_col = headers.get("VESSEL")
        branch_col = headers.get("BRANCH")
        etd_col = headers.get("ETD")

        total_rows = ws.max_row
        for row_idx in range(2, total_rows + 1):
            bk_value = ws.cell(row=row_idx, column=bk_col).value

            if not bk_value or str(bk_value).strip() == "":
                # BK # trống => bỏ qua
                continue

            bk_value = str(bk_value).strip()

            # Truy vấn trong master data
            cursor.execute(f"SELECT * FROM {TABLE_NAME} WHERE `Carrier Bkg no.` = ?", (bk_value,))
            rows = cursor.fetchall()

            if len(rows) == 0:
                print(f"⚠️ BK # : {bk_value} không tìm thấy trong master data.")
                continue
            elif len(rows) > 1:
                print(f"⚠️ BK # : {bk_value} đang bị trùng lặp trong master data, phát hiện {len(rows)} dòng. Sử dụng dòng đầu tiên.")

            row = rows[0]
            col_names = [desc[0] for desc in cursor.description]
            data = dict(zip(col_names, row))

            # Tính toán ETD
            etd_value = ""
            origin_etd = data.get("Origin ETD", "").strip()
            new_etd = data.get("New ETD", "").strip()

            if origin_etd and not new_etd:
                etd_value = origin_etd
            elif origin_etd and new_etd:
                etd_value = new_etd

            # Cập nhật từng ô tương ứng
            if line_col: ws.cell(row=row_idx, column=line_col).value = data.get("Carrier", "")
            if job_col: ws.cell(row=row_idx, column=job_col).value = data.get("Job no.", "")
            if vessel_col: ws.cell(row=row_idx, column=vessel_col).value = data.get("Vessel name", "")
            if branch_col: ws.cell(row=row_idx, column=branch_col).value = data.get("Shipper Name", "")
            if etd_col: ws.cell(row=row_idx, column=etd_col).value = etd_value

        # Lưu lại file gốc
        wb.save(file_path)
        conn.close()
        print(f"✅ Hoàn tất fill data và ghi đè vào file gốc: {file_path}")

    except KeyboardInterrupt:
        print("\n🛑 Quá trình bị hủy bởi người dùng (Ctrl + C). Quay lại menu chính...")
    except Exception as e:
        print(f"❌ Lỗi trong quá trình fill data: {e}")



def main():
    try:
        while True:
            choice = show_menu()
            if choice == "1":
                init_master_data()
            elif choice == "2":
                fill_data()
            elif choice == "0":
                print("👋 Thoát chương trình. Tạm biệt!")
                break
            else:
                print("❌ Lựa chọn không hợp lệ, vui lòng nhập lại.")
    except KeyboardInterrupt:
        print("\n👋 Thoát chương trình bằng Ctrl + C. Tạm biệt!")


if __name__ == "__main__":
    main()
