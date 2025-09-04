## 1. Tải toàn bộ src code và đặt trong một thư mục riêng.

File `shipping_classifier.py` sẽ nằm trong thư mục đó.

## 2. Chuẩn bị dữ liệu đầu vào
Tạo file `input.csv` chứa dữ liệu cần xử lý.

Đặt `input.csv` cùng cấp với `shipping_classifier.py`.

Ví dụ cấu trúc thư mục:
```json
shipping-classifier/
├── shipping_classifier.py
├── input.csv
```
## 3. Kiểm tra phiên bản Python và cài đặt pandas
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
python3 -m pip install pandas
```
Lưu ý: Nếu bạn dùng Python 2, hãy thay đổi lệnh trên thành python -m pip install pandas.

## 4. Chạy chương trình
Trong thư mục chứa `shipping_classifier.py` và `input.csv`, chạy lệnh:

 Thay thế 'python3' bằng lệnh bạn tìm được
```json
python3 shipping_classifier.py
```
## 5. Kiểm tra kết quả
Sau khi chạy xong, kết quả sẽ được lưu trong output.txt:


```json
shipping-classifier/
├── shipping_classifier.py
├── input.csv
└── output.txt
```
Mở file output.txt để xem kết quả.

## Note : Lưu ý kiểm tra kỹ kết quả để tránh có sai sót phát sinh

