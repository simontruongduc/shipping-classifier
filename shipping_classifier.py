# -*- coding: utf-8 -*-
import pandas as pd
import os
import sys

def process_csv(input_file, output_file):
    try:
        # Đọc dữ liệu CSV
        df = pd.read_csv(input_file)

        # Các cột cần thiết
        required_cols = ["Container NO", "UN Type", "HTS CODE", "Desc.of Goods"]
        for col in required_cols:
            if col not in df.columns:
                print("❌ Lỗi: Không tìm thấy cột '{}' trong file CSV.".format(col))
                sys.exit(1)

        # Bỏ các dòng UN Type rỗng
        df = df[df["UN Type"].notna()]
        df = df[df["UN Type"].astype(str).str.strip() != ""]

        # Fill giá trị Container NO trống
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
        with open(output_file, "w", encoding="utf-8") if sys.version_info.major >= 3 else open(output_file, "w") as f:
            # Loop từng container
            for container_no, rows in grouped.items():
                container_df = pd.DataFrame(rows)

                # Loại bỏ trùng theo primary key
                container_df = container_df.drop_duplicates(
                    subset=["UN Type", "HTS CODE", "Desc.of Goods"]
                )

                # Phân loại theo UN Type
                dg_df = container_df[container_df["UN Type"].str.upper() == "DG"]
                nondg_df = container_df[container_df["UN Type"].str.upper() == "NONDG"]
                general_df = container_df[container_df["UN Type"].str.upper() == "GENERAL"]

                # --- Xử lý DG ---
                if not dg_df.empty:
                    f.write("cont: {}\n".format(container_no))
                    for desc in dg_df["Desc.of Goods"].dropna().tolist():
                        f.write("{}\n".format(desc))
                    f.write("hs code: {}\n".format(", ".join(dg_df["HTS CODE"].dropna().astype(str).tolist())))
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
                    f.write("hs code: {}\n".format(", ".join(nondg_df["HTS CODE"].dropna().astype(str).tolist())))
                    f.write("NONDG GOODS CONTAIN BATTERY\n")
                    f.write("---------------------------------\n")

                # --- Xử lý GENERAL ---
                if not general_df.empty:
                    f.write("cont: {}\n".format(container_no))
                    for desc in general_df["Desc.of Goods"].dropna().tolist():
                        f.write("{}\n".format(desc))
                    f.write("hs code: {}\n".format(", ".join(general_df["HTS CODE"].dropna().astype(str).tolist())))
                    f.write("GENERAL GOODS WITHOUT BATTERY\n")
                    f.write("---------------------------------\n")

        print("✅ Đã xử lý xong! File xuất: {}".format(output_file))

    except Exception as e:
        print("❌ Lỗi: {}".format(e))


if __name__ == "__main__":
    # Đường dẫn file input.csv và output.txt cùng thư mục với file .py
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_file = os.path.join(base_dir, "input.csv")
    output_file = os.path.join(base_dir, "output.txt")

    process_csv(input_file, output_file)
