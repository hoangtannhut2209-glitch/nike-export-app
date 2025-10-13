# Nike Export Web App

Hệ thống web app xử lý chứng từ xuất khẩu Nike tự động từ PDF sang Excel templates.

## 🚀 Tính năng chính

- **PDF Processing**: Tự động trích xuất dữ liệu từ PL và BOOKING PDF files
- **Template Management**: Quản lý templates Excel với hệ thống placeholder
- **Batch Processing**: Xử lý nhiều invoice cùng lúc
- **Web Interface**: Giao diện web hiện đại với Bootstrap
- **API Support**: RESTful APIs cho integration
- **Database Storage**: Lưu trữ dữ liệu với SQLite

## 📋 Yêu cầu hệ thống

- Python 3.8+
- pip
- Virtual environment (khuyên dùng)

## 🛠️ Cài đặt

### Tự động (khuyên dùng)

```bash
chmod +x setup.sh
./setup.sh
```

### Thủ công

1. **Tạo virtual environment:**
```bash
python3 -m venv venv
source venv/bin/activate  # Linux/Mac
# hoặc
venv\Scripts\activate     # Windows
```

2. **Cài đặt dependencies:**
```bash
pip install -r requirements.txt
```

3. **Tạo thư mục cần thiết:**
```bash
mkdir -p uploads outputs excel_templates
```

4. **Khởi tạo database:**
```bash
python database.py
```

5. **Tạo sample templates:**
```bash
python -c "from core.template_manager import create_sample_template; create_sample_template('excel_templates/sample.xlsx')"
```

## 🏃‍♂️ Chạy ứng dụng

```bash
source venv/bin/activate
python app.py
```

Truy cập: http://localhost:5000

## 📖 Cách sử dụng

### 1. Upload Files

1. Vào trang **Upload Files**
2. Nhập Invoice Number
3. Chọn PL PDF file và/hoặc BOOKING PDF file
4. Click **Upload và Xử lý**

### 2. Quản lý Templates

1. Vào trang **Templates**
2. Upload template Excel với placeholders như `{Invoice Number}`, `{Date}`, etc.
3. Hoặc sử dụng sample templates có sẵn

### 3. Tạo chứng từ

1. Vào **Templates** → **Generate Template**
2. Chọn template và invoice data
3. Download file Excel đã được điền dữ liệu

## 🔧 Cấu hình

### Placeholders hỗ trợ

| Placeholder | Mô tả |
|-------------|--------|
| `{Invoice Number}` | Số invoice |
| `{Date}` | Ngày (dd/mm/yyyy) |
| `{PO}` | Purchase Order number |
| `{Reference PO#}` | Reference PO number |
| `{Material}` | Material code |
| `{DESCRIPTION}` | Mô tả sản phẩm từ BOOKING |
| `{Marks}` | Marks (IN HOA) |
| `{DEST}` | Destination country |
| `{Total Cartons}` | Tổng số thùng |
| `{Total Cartons In Words}` | Số thùng bằng chữ |
| `{Plant}` | Plant code |
| `{Customer Ship To #}` | Customer Ship To number |
| `{Total Gross Kgs}` | Tổng trọng lượng |
| `{Total Units}` | Tổng số lượng |

### File types hỗ trợ

- **PDF**: .pdf (cho PL và BOOKING files)
- **Excel**: .xlsx, .xlsm, .xls (cho templates)

## 🏗️ Cấu trúc project

```
nike-export-webapp/
├── app.py                 # Main Flask app
├── config.py              # Configuration
├── database.py            # Database models
├── requirements.txt       # Dependencies
├── setup.sh              # Setup script
├── core/                 # Core modules
│   ├── pdf_processor.py  # PDF processing logic
│   └── template_manager.py # Template management
├── routes/               # Flask routes
│   ├── main.py          # Main routes
│   ├── api.py           # API endpoints
│   └── templates.py     # Template routes
├── templates/            # HTML templates
│   ├── base.html
│   ├── index.html
│   ├── upload.html
│   └── templates/
├── static/               # Static files
│   ├── css/
│   └── js/
├── uploads/              # Uploaded files
├── outputs/              # Generated files
└── excel_templates/      # Template files
```

## 🔄 API Endpoints

### Upload & Processing
- `POST /api/upload` - Upload và xử lý files
- `GET /api/job/{job_id}/status` - Kiểm tra tiến trình xử lý
- `POST /api/process/{invoice_number}` - Xử lý invoice cụ thể

### Invoices
- `GET /api/invoices` - List tất cả invoices
- `GET /api/invoices/{invoice_number}` - Chi tiết invoice

### Templates
- `POST /api/templates/generate` - Tạo template đã điền dữ liệu
- `POST /api/templates/upload` - Upload template mới
- `GET /api/templates/{template_name}/placeholders` - Lấy danh sách placeholders
- `POST /api/templates/batch-generate` - Tạo batch templates

### Dashboard
- `GET /api/dashboard/stats` - Thống kê dashboard
- `GET /api/dashboard/recent-activity` - Hoạt động gần đây

## 🐛 Troubleshooting

### Lỗi import

Nếu gặp lỗi import, đảm bảo đã activate virtual environment:
```bash
source venv/bin/activate
```

### Lỗi PDF processing

- Kiểm tra file PDF không bị corrupt
- Đảm bảo file có text layer (không phải scan)
- Kiểm tra format file đúng (PL hoặc BOOKING)

### Lỗi template

- Kiểm tra file Excel không bị password protect
- Đảm bảo placeholders sử dụng đúng format `{Field Name}`
- Kiểm tra file không quá lớn (< 50MB)

## 📞 Hỗ trợ

- Kiểm tra logs trong terminal khi chạy app
- Sử dụng browser developer tools để debug frontend
- Check database với SQLite browser

## 📄 License

MIT License - Sử dụng tự do cho mục đích thương mại và cá nhân.

## 🔄 Updates

Version 1.0.0 - Initial release
- Basic PDF processing
- Template management
- Web interface
- API endpoints