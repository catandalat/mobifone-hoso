# Hồ sơ tiếp khách – MobiFone Lâm Đồng

Web app tạo tự động 4 file Word cho bộ hồ sơ thanh toán tiếp khách.

## Các file được tạo
1. `Tờ Trình` – trình lãnh đạo duyệt chi phí
2. `Giấy đề nghị tiếp khách` – mẫu 08-TT
3. `Bảng kê thanh toán` – bảng kê CP + bảng kê khai thường xuyên
4. `Báo cáo kết quả công việc`

## Cấu trúc thư mục
```
mobifone-app/
├── app.py                  # Flask backend
├── utils.py                # Tiện ích (số tiền bằng chữ, format ngày)
├── build_templates.py      # Script tạo file Word template
├── requirements.txt
├── Procfile
├── railway.json
├── templates/
│   └── index.html          # Giao diện web
└── word_templates/         # 4 file .docx template (auto-generated)
    ├── to_trinh.docx
    ├── giay_de_nghi.docx
    ├── bang_ke.docx
    └── bao_cao_kqcv.docx
```

## Deploy lên Railway

### Bước 1: Cài Railway CLI
```bash
npm install -g @railway/cli
```

### Bước 2: Đăng nhập và tạo project
```bash
railway login
railway init
```

### Bước 3: Deploy
```bash
railway up
```

### Bước 4: Lấy domain
```bash
railway domain
```

Railway tự động detect Python, cài requirements.txt và chạy Procfile.

## Chạy local (để test)
```bash
pip install -r requirements.txt
python build_templates.py   # Tạo file template (chỉ cần chạy 1 lần)
python app.py
```
Mở http://localhost:5000

## Cập nhật nội dung
- **Thay đổi người ký mặc định**: sửa value trong `templates/index.html`
- **Thay đổi layout văn bản**: sửa `build_templates.py` rồi chạy lại
- **Thêm loại hồ sơ mới**: thêm template mới vào `word_templates/` và đăng ký trong `app.py`
