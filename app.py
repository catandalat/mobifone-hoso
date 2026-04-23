from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from docxtpl import DocxTemplate
import io
import os
import zipfile
import json
import base64
import re
from datetime import datetime
from utils import so_tien_bang_chu, format_date, format_currency

# ── Claude API cho đọc hóa đơn ───────────────────────────────────────────────
try:
    import anthropic
    _claude = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))
except Exception:
    _claude = None

try:
    import pdfplumber
except Exception:
    pdfplumber = None

app = Flask(__name__)
CORS(app)

TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "word_templates")

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/read-invoice", methods=["POST"])
def read_invoice():
    """
    Nhận PDF hóa đơn, trích text bằng pdfplumber,
    gửi Claude API để parse JSON các trường cần thiết.
    """
    if "file" not in request.files:
        return jsonify({"error": "Không có file PDF"}), 400

    pdf_file = request.files["file"]
    if not pdf_file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Vui lòng upload file PDF"}), 400

    # ── Bước 1: Trích text từ PDF ────────────────────────────────────────────
    try:
        pdf_bytes = pdf_file.read()
        text = ""
        if pdfplumber:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text() or ""
                    text += page_text + "\n"
        if not text.strip():
            return jsonify({"error": "Không đọc được nội dung PDF. Thử lại với file khác."}), 400
    except Exception as e:
        return jsonify({"error": f"Lỗi đọc PDF: {str(e)}"}), 500

    # ── Bước 2: Claude API parse JSON ────────────────────────────────────────
    if not _claude or not os.environ.get("ANTHROPIC_API_KEY"):
        return jsonify({"error": "Chưa cấu hình ANTHROPIC_API_KEY"}), 500

    prompt = f"""Bạn là trợ lý đọc hóa đơn điện tử Việt Nam. 
Hãy đọc nội dung hóa đơn dưới đây và trích xuất các trường thông tin.
Trả về DUY NHẤT một JSON object, không có text nào khác, không có markdown.

Các trường cần trích xuất:
- so_hd: số hóa đơn (chỉ số, bỏ số 0 đầu — ví dụ "00001739" → "1739")
- ky_hieu_hd: ký hiệu hóa đơn (ví dụ "1C26MNC")
- ngay_hd: ngày hóa đơn định dạng YYYY-MM-DD (ví dụ "2026-03-27")
- nha_cung_cap: tên công ty bán hàng (tên đầy đủ)
- mst_ncc: mã số thuế người bán
- dc_ncc: địa chỉ người bán (rút gọn, bỏ "Số" đầu nếu có)
- truoc_vat: tổng tiền trước VAT (chỉ số nguyên, không dấu chấm/phẩy)
- tien_vat: tổng tiền VAT (chỉ số nguyên)
- sau_vat: tổng tiền sau VAT (chỉ số nguyên)
- thang_tt: tháng thanh toán dạng MM/YYYY (lấy từ ngày hóa đơn)

Lưu ý:
- Nếu hóa đơn có nhiều mức thuế suất, cộng tất cả lại
- truoc_vat + tien_vat = sau_vat
- Nếu không tìm thấy trường nào, để chuỗi rỗng ""
- Chỉ trả về JSON thuần, không giải thích

NỘI DUNG HÓA ĐƠN:
{text[:3000]}
"""

    try:
        resp = _claude.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=800,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = resp.content[0].text.strip()

        # Lấy JSON từ response (phòng trường hợp có text thừa)
        json_match = re.search(r'\{.*\}', raw, re.DOTALL)
        if not json_match:
            return jsonify({"error": "Claude không trả về JSON hợp lệ"}), 500

        data = json.loads(json_match.group())

        # Validate & clean số tiền
        for key in ["truoc_vat", "tien_vat", "sau_vat"]:
            val = str(data.get(key, "0")).replace(".", "").replace(",", "").strip()
            data[key] = int(val) if val.isdigit() else 0

        return jsonify({"ok": True, "data": data})

    except json.JSONDecodeError as e:
        return jsonify({"error": f"Lỗi parse JSON: {str(e)}", "raw": raw}), 500
    except Exception as e:
        return jsonify({"error": f"Lỗi Claude API: {str(e)}"}), 500


@app.route("/api/generate", methods=["POST"])
def generate_docs():
    data = request.json
    if not data:
        return jsonify({"error": "Không có dữ liệu"}), 400

    # Chuẩn bị context chung
    truoc_vat = int(data.get("truoc_vat", 0))
    tien_vat  = int(data.get("tien_vat", 0))
    sau_vat   = truoc_vat + tien_vat
    sl_ld     = int(data.get("sl_ld", 0))
    sl_cv     = int(data.get("sl_cv", 0))
    sl_khach  = int(data.get("sl_khach", 0))

    ctx = {
        # Người đề nghị
        "ho_ten":           data.get("ho_ten", ""),
        "don_vi":           data.get("don_vi", ""),
        "lanh_dao":         data.get("lanh_dao", ""),
        "chuc_danh_ld":     data.get("chuc_danh_ld", ""),
        "giam_doc":         data.get("giam_doc", ""),
        "phu_trach_cp":     data.get("phu_trach_cp", ""),
        # Thời gian
        "ngay_tiep_khach":  format_date(data.get("ngay_tiep_khach", "")),
        "ngay_to_trinh":    format_date(data.get("ngay_to_trinh", "")),
        "ngay_bao_cao":     format_date(data.get("ngay_bao_cao", "")),
        "ngay_hd":          format_date(data.get("ngay_hd", "")),
        "thang_tt":         data.get("thang_tt", ""),
        "so_to_trinh":      data.get("so_to_trinh", ""),
        "ma_kmcp":          data.get("ma_kmcp", ""),
        # Hóa đơn
        "so_hd":            data.get("so_hd", ""),
        "ky_hieu_hd":       data.get("ky_hieu_hd", ""),
        "nha_cung_cap":     data.get("nha_cung_cap", ""),
        "mst_ncc":          data.get("mst_ncc", ""),
        "dc_ncc":           data.get("dc_ncc", ""),
        # Tiền
        "truoc_vat":        format_currency(truoc_vat),
        "tien_vat":         format_currency(tien_vat),
        "sau_vat":          format_currency(sau_vat),
        "truoc_vat_raw":    truoc_vat,
        "tien_vat_raw":     tien_vat,
        "sau_vat_raw":      sau_vat,
        "tien_bang_chu":    so_tien_bang_chu(truoc_vat),
        "tong_bang_chu":    so_tien_bang_chu(sau_vat),
        # Kế toán
        "tk_kt":            data.get("tk_kt", ""),
        "nghiep_vu":        data.get("nghiep_vu", ""),
        "ma_spdv":          data.get("ma_spdv", ""),
        "quyet_dinh_cp":    data.get("quyet_dinh_cp", ""),
        # Nội dung tiếp khách
        "khach_moi":        data.get("khach_moi", ""),
        "ly_do":            data.get("ly_do", ""),
        "ket_qua":          data.get("ket_qua", ""),
        "sl_ld":            sl_ld,
        "sl_cv":            sl_cv,
        "sl_khach":         sl_khach,
        "tong_nguoi":       sl_ld + sl_cv + sl_khach,
        # Ngày dạng số riêng (cho các văn bản cần dd/mm/yyyy)
        "ngay_tk_so":       _day(data.get("ngay_tiep_khach", "")),
        "thang_tk_so":      _month(data.get("ngay_tiep_khach", "")),
        "nam_tk_so":        _year(data.get("ngay_tiep_khach", "")),
        "ngay_tt_so":       _day(data.get("ngay_to_trinh", "")),
        "thang_tt_so":      _month(data.get("ngay_to_trinh", "")),
        "nam_tt_so":        _year(data.get("ngay_to_trinh", "")),
        "ngay_bc_so":       _day(data.get("ngay_bao_cao", "")),
        "thang_bc_so":      _month(data.get("ngay_bao_cao", "")),
        "nam_bc_so":        _year(data.get("ngay_bao_cao", "")),
    }

    # Danh sách file cần tạo
    templates = [
        ("to_trinh.docx",        "2_TờTrình_TiếpKhách.docx"),
        ("giay_de_nghi.docx",    "3_GiấyĐềNghịTiếpKhách.docx"),
        ("bang_ke.docx",         "4_BảngKê.docx"),
        ("bao_cao_kqcv.docx",    "5_BáoCáoKQCV.docx"),
    ]

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for tpl_name, out_name in templates:
            tpl_path = os.path.join(TEMPLATE_DIR, tpl_name)
            if not os.path.exists(tpl_path):
                continue
            doc = DocxTemplate(tpl_path)
            doc.render(ctx)
            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            zf.writestr(out_name, doc_buffer.read())

    zip_buffer.seek(0)
    filename = f"HoSoTiepKhach_{data.get('thang_tt','').replace('/','_')}.zip"
    return send_file(
        zip_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name=filename
    )

def _day(date_str):
    try:
        return str(datetime.strptime(date_str, "%Y-%m-%d").day).zfill(2)
    except Exception:
        return ""

def _month(date_str):
    try:
        return str(datetime.strptime(date_str, "%Y-%m-%d").month).zfill(2)
    except Exception:
        return ""

def _year(date_str):
    try:
        return str(datetime.strptime(date_str, "%Y-%m-%d").year)
    except Exception:
        return ""

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
