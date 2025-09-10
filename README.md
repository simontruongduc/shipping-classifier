## 1. Tải toàn bộ src code và đặt trong một thư mục riêng.

File `shipping_classifier.py` sẽ nằm trong thư mục đó.

## 2. Kiểm tra phiên bản Python và cài đặt pandas
Để cài đặt thư viện pandas, trước hết bạn cần biết máy tính của bạn sử dụng lệnh nào để chạy Python. Hãy thử từng lệnh dưới đây:

### a. Kiểm tra lệnh chạy Python

Mở terminal (trên macOS/Linux) hoặc Command Prompt (trên Windows) và chạy lần lượt các lệnh sau cho đến khi một lệnh hiển thị phiên bản Python của bạn:


```json
python --version
python3 --version
py --version
```
### b. Cài đặt pandas

Sau khi tìm được lệnh đúng (ví dụ: python3), hãy dùng lệnh đó để cài đặt pandas:

Thay thế 'python3' bằng lệnh bạn tìm được ở trên
```json
python3 -m pip install pandas openpyxl xlrd
```
Lưu ý: Nếu bạn dùng Python 2, hãy thay đổi lệnh trên thành python -m pip install pandas.

## 4. Chạy chương trình
Trong thư mục chứa `shipping_classifier.py` chạy lệnh:

 Thay thế 'python3' bằng lệnh bạn tìm được
```json
python3 shipping_classifier.py
```
kéo thả file input vào 
nếu là file excel ( xls , xlsx ) , chọn sheet cần xử lý

## 5. Kiểm tra kết quả
Sau khi chạy xong, kết quả sẽ được lưu trong _output.txt

Mở file _output.txt để xem kết quả.

## Note : Lưu ý kiểm tra kỹ kết quả để tránh có sai sót phát sinh

