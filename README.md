# 🧾 Auto Scan PO - Customs Classification Assistant

**Mục tiêu:**  
Tự động hóa quá trình xử lý email PO, phân loại lô hàng có cần khai báo hải quan (CDs) hay không, và gửi yêu cầu bổ sung thông tin cho supplier phục vụ khai báo HS code.

---

## 📁 Modules chính

| Module                      | Mô tả chức năng chính                                                                 |
|-----------------------------|---------------------------------------------------------------------------------------|
| `m01_email_reader.py`       | Đọc email từ Outlook, tải file đính kèm PDF từ thư mục cấu hình                       |
| `m02_pdf_scan.py`           | Phân tích nội dung file PDF, trích xuất thông tin PO và xác định có cần CDs không     |
| `m03_send_request_email.py` | Tự động gửi email yêu cầu cung cấp thông tin hàng hóa cho các PO cần khai báo hải quan|
| `gui_main.py`               | Giao diện người dùng (GUI) cho phép chọn thư mục, nhập config, scan email & gửi mail  |
| `config.py`                 | Cấu hình tập trung theo class `Settings` dễ tùy biến và mở rộng                       |
| `utils.py`                  | Hàm phụ trợ dùng chung, ví dụ: resolve email Exchange                                 |

---

## 🧠 Logic phân loại cần CDs (customs declaration sheet)

Xác định theo luồng logic trong `classify logic.txt`:

```python
if Currency != "VND":
    Need_CDs = "Yes"
else:
    if "0%" not in VAT:
        Need_CDs = "No"
    elif UOM == "UNIT":
        Need_CDs = "No"
    elif Seller in Non-CDs Supplier List:
        Need_CDs = "No"
    elif Max Unit Price > 30.000.000:
        Need_CDs = "Yes"
    elif UOM không chứa PIECE hoặc SET:
        Need_CDs = "No"
    else:
        Need_CDs = "Yes"
```
---
## 📁 Cấu trúc thư mục
```bash
├── .gitignore
├── README.md
├── requirement.txt
├── 0_Run_Files/
│   ├── config.py
│   ├── gui_main.py
│   ├── m01_email_reader.py
│   ├── m02_pdf_scan.py
│   ├── m03_send_request_email.py
│   ├── utils.py
├── temp
│   ├── 1_LOCAL HS code request.xlsx    # Template file Cargo info cho hàng Local
│   ├── 2_OVERSEA Machine list.xlsx     # Template file Cargo info cho hàng Oversea
│   ├── classify logic.txt              # Logic Classify PO Need CDs
│   ├── Non-CDs Supplier.csv            # List of Non-CDs Suppliers (Service, Office Supply, Safety Workwear,...) 

```
---

## 🗂 Cấu trúc thư mục đầu ra

```bash
Scanned PO/
├── temp/                   # Lưu file PDF tải về tạm thời
├── PO_Filtered/            # Các file PDF phân loại cần CDs, chia theo Buyer
│   ├── 1. TTIVN MFG/
│   ├── 2. GREEN PLANET/
│   └── ...
├── log/
│   ├── po_log.csv          # Tổng hợp kết quả phân loại
│   ├── thread_*.csv        # Log theo luồng xử lý song song
│   └── error.txt           # Ghi lỗi khi xử lý PDF
```

---

## ▶️ Cách chạy

### CLI:
```bash
python m01_email_reader.py        # Đọc email & tải PDF
python m02_pdf_scan.py            # Phân tích PDF & xác định cần CDs
python m03_send_request_email.py  # Gửi email yêu cầu cung cấp thông tin
```

### GUI:
```bash
python gui_main.py
```

---

## 📌 Yêu cầu hệ thống

- Windows + Outlook Desktop
- Python >= 3.10
- Thư viện: `pandas`, `pdfplumber`, `jinja2`, `tkcalendar`, `pywin32`

---

## 📄 License

Private internal tool – Customs & Compliance team – TTI Vietnam.