# -*- coding: utf-8 -*-
import subprocess
import sys

def show_menu():
    print("\n=== MAIN MENU ===")
    print("1. Auto Fill")
    print("2. Shipping Classifier")
    print("0. Thoát")
    return input("Nhập lựa chọn của bạn: ").strip()

def run_script(script_name):
    try:
        print(f"\n▶️ Đang chạy chương trình: {script_name} ...\n")
        subprocess.run([sys.executable, script_name], check=True)  # ← Dùng chính Python đang chạy main.py
    except subprocess.CalledProcessError as e:
        print(f"❌ Lỗi khi chạy {script_name}: {e}")
    except KeyboardInterrupt:
        print("\n👋 Dừng chương trình.")

def main():
    try:
        while True:
            choice = show_menu()
            if choice == "1":
                run_script("auto_fill.py")
            elif choice == "2":
                run_script("shipping_classifier.py")
            elif choice == "0":
                print("👋 Thoát chương trình.")
                break
            else:
                print("❌ Lựa chọn không hợp lệ, vui lòng nhập lại.")
    except KeyboardInterrupt:
        print("\n👋 Tạm biệt!")

if __name__ == "__main__":
    main()
