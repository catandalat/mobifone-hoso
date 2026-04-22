"""
Tạo 4 file Word template cho bộ hồ sơ tiếp khách MobiFone Lâm Đồng.
Chạy: python build_templates.py
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

OUT = "word_templates"
os.makedirs(OUT, exist_ok=True)

FONT = "Times New Roman"
PAGE_W = Cm(21)   # A4
PAGE_H = Cm(29.7)
MARGIN = Cm(2.5)

def set_margins(doc, top=2.5, bottom=2.5, left=3, right=2):
    for section in doc.sections:
        section.top_margin    = Cm(top)
        section.bottom_margin = Cm(bottom)
        section.left_margin   = Cm(left)
        section.right_margin  = Cm(right)

def para(doc, text="", bold=False, italic=False, size=13, align=WD_ALIGN_PARAGRAPH.LEFT,
         space_before=0, space_after=6, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    p.alignment = align
    run = p.add_run(text)
    run.bold   = bold
    run.italic = italic
    run.font.name = FONT
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)
    return p

def two_col_header(doc, left_lines, right_lines, size=12):
    """Header 2 cột: trái = đơn vị, phải = quốc hiệu"""
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.style = "Table Grid"
    # Xóa border
    for cell in tbl.rows[0].cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        for side in ["top","left","bottom","right","insideH","insideV"]:
            b = OxmlElement(f"w:{side}")
            b.set(qn("w:val"), "none")
            tcBorders.append(b)
        tcPr.append(tcBorders)

    def fill_cell(cell, lines, align=WD_ALIGN_PARAGRAPH.CENTER):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        for i, (txt, bold_, size_) in enumerate(lines):
            if i == 0 and not cell.paragraphs[0].runs:
                p = cell.paragraphs[0]
            else:
                p = cell.add_paragraph()
            p.alignment = align
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(2)
            run = p.add_run(txt)
            run.bold = bold_
            run.font.name = FONT
            run.font.size = Pt(size_)

    fill_cell(tbl.rows[0].cells[0], left_lines, WD_ALIGN_PARAGRAPH.CENTER)
    fill_cell(tbl.rows[0].cells[1], right_lines, WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)

def add_table_border(tbl):
    tblPr = tbl._tbl.tblPr
    tblBorders = OxmlElement("w:tblBorders")
    for side in ["top","left","bottom","right","insideH","insideV"]:
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), "000000")
        tblBorders.append(b)
    tblPr.append(tblBorders)

def cell_text(cell, text, bold=False, size=12, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(1)
    run = p.add_run(text)
    run.bold = bold
    run.font.name = FONT
    run.font.size = Pt(size)

# ─────────────────────────────────────────────
# 1. TỜ TRÌNH
# ─────────────────────────────────────────────
def make_to_trinh():
    doc = Document()
    set_margins(doc)

    two_col_header(doc,
        left_lines=[
            ("TCT VIỄN THÔNG MOBIFONE", False, 11),
            ("MOBIFONE LÂM ĐỒNG", True,  11),
            ("TT KD GIẢI PHÁP SỐ",   True,  11),
            ("───────────────────", False, 9),
            ("Số: {{so_to_trinh}}/TTr-TTKDGPS", False, 11),
        ],
        right_lines=[
            ("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", True, 11),
            ("Độc lập – Tự do – Hạnh phúc", False, 11),
            ("───────────────────", False, 9),
            ("Lâm Đồng, {{ngay_to_trinh}}", False, 11),
        ]
    )

    para(doc, "TỜ TRÌNH", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    para(doc, "Về việc: Đề xuất chi phí tiếp khách tại MobiFone Lâm Đồng tháng {{thang_tt}}",
         bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    para(doc, "Kính trình: Ông {{giam_doc}} – Giám đốc MobiFone Lâm Đồng",
         size=13, space_after=6)

    # Ô ý kiến chỉ đạo
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    add_table_border(tbl)
    cell_text(tbl.rows[0].cells[0],
              "Ý kiến chỉ đạo của Lãnh đạo:\n\n\n",
              bold=True, size=12)

    para(doc, "", space_after=4)
    para(doc,
         "Căn cứ Quyết định số {{quyet_dinh_cp}} về việc giao kế hoạch chi phí lần 1 năm {{nam_tt_so}};",
         size=13, space_after=4)
    para(doc, "Căn cứ nhu cầu thực tế tại đơn vị.", size=13, space_after=6)
    para(doc,
         "Trung tâm Kinh doanh Giải pháp số kính trình Lãnh đạo về việc duyệt chi phí tiếp khách "
         "tại MobiFone Lâm Đồng, cụ thể như sau:",
         size=13, space_after=6)

    para(doc, "- Mục tiêu, lý do:", bold=True, size=13, space_after=2)
    para(doc, "  Nhằm đảm bảo sản xuất kinh doanh tại MobiFone Lâm Đồng.", size=13, space_after=2)
    para(doc, "  {{ly_do}}", size=13, space_after=6)

    para(doc, "- Nội dung thực hiện:", bold=True, size=13, space_after=2)
    para(doc, "  Mời cơm thân mật cùng với các Doanh nghiệp trên địa bàn.", size=13, space_after=2)
    para(doc, "  Trao đổi, giải đáp thắc mắc của doanh nghiệp đối với các sản phẩm của MobiFone.",
         size=13, space_after=2)
    para(doc, "  Thành phần tham dự: Lãnh đạo, chuyên viên liên quan, đại diện các Doanh nghiệp.",
         size=13, space_after=6)

    para(doc, "- Chi phí dự kiến: {{truoc_vat}} đồng (chưa bao gồm VAT). Bằng chữ: {{tien_bang_chu}}.",
         bold=False, size=13, space_after=4)
    para(doc, "- Nguồn chi phí:", bold=False, size=13, space_after=2)
    para(doc, "  Quyết định số {{quyet_dinh_cp}}.", size=13, space_after=2)
    para(doc, "  KMCP: {{ma_kmcp}} – MOBIFONE - VT - Di động - CPQLDN - CP Tiếp khách, khánh tiết - Quản lý hành chính",
         size=13, space_after=4)
    para(doc, "- Thời gian thực hiện: Tháng {{thang_tt}}", size=13, space_after=6)

    para(doc,
         "Kính trình Lãnh đạo xem xét và phê duyệt cho Trung tâm Kinh doanh Giải pháp số được "
         "chủ trì bố trí kế hoạch cụ thể trên địa bàn nhằm phù hợp với thực tế./.",
         size=13, space_after=10)

    # Chữ ký
    tbl2 = doc.add_table(rows=2, cols=3)
    tbl2.style = "Table Grid"
    for row in tbl2.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcB = OxmlElement("w:tcBorders")
            for side in ["top","left","bottom","right"]:
                b = OxmlElement(f"w:{side}")
                b.set(qn("w:val"), "none")
                tcB.append(b)
            tcPr.append(tcB)

    cell_text(tbl2.rows[0].cells[0], "Nơi nhận:", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT)
    cell_text(tbl2.rows[0].cells[1], "", size=11)
    cell_text(tbl2.rows[0].cells[2], "PHỤ TRÁCH TRUNG TÂM", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl2.rows[1].cells[0],
              "- Như trên;\n- Lưu P.TH.", size=11, align=WD_ALIGN_PARAGRAPH.LEFT)
    cell_text(tbl2.rows[1].cells[1], "", size=11)
    cell_text(tbl2.rows[1].cells[2], "\n\n{{lanh_dao}}", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.save(f"{OUT}/to_trinh.docx")
    print("✓ to_trinh.docx")

# ─────────────────────────────────────────────
# 2. GIẤY ĐỀ NGHỊ TIẾP KHÁCH
# ─────────────────────────────────────────────
def make_giay_de_nghi():
    doc = Document()
    set_margins(doc)

    two_col_header(doc,
        left_lines=[
            ("TCT VIỄN THÔNG MOBIFONE", False, 11),
            ("MOBIFONE LÂM ĐỒNG", True, 11),
            ("───────────────────", False, 9),
            ("MẪU SỐ: 08-TT", False, 11),
        ],
        right_lines=[
            ("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", True, 11),
            ("Độc lập – Tự do – Hạnh phúc", False, 11),
        ]
    )

    para(doc, "GIẤY ĐỀ NGHỊ TIẾP KHÁCH", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, "{{ngay_tiep_khach}}", italic=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)

    para(doc, "Kính gửi: Lãnh đạo MobiFone Lâm Đồng", bold=True, size=13, space_after=6)
    para(doc, "\tĐơn vị đề nghị: {{don_vi}}", size=13, space_after=4)
    para(doc, "Lý do: {{ly_do}}", size=13, space_after=4)
    para(doc, "\tNgười chủ trì: {{ho_ten}}", size=13, space_after=4)
    para(doc, "\tThành phần và số người tham gia", size=13, space_after=2)
    para(doc, "\t+ Về phía MBF Lâm Đồng: {{sl_ld}} Lãnh đạo và {{sl_cv}} Chuyên viên", size=13, space_after=2)
    para(doc, "\t+ Về phía khách mời: {{sl_khach}} người", size=13, space_after=4)
    para(doc, "\tKhách mời: {{khach_moi}}", size=13, space_after=4)
    para(doc,
         "\tDự trù số tiền chi tiếp khách: {{truoc_vat}} đồng (chưa VAT). Bằng chữ: {{tien_bang_chu}}.",
         size=13, space_after=10)

    tbl = doc.add_table(rows=2, cols=2)
    for row in tbl.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcB = OxmlElement("w:tcBorders")
            for side in ["top","left","bottom","right"]:
                b = OxmlElement(f"w:{side}")
                b.set(qn("w:val"), "none")
                tcB.append(b)
            tcPr.append(tcB)

    cell_text(tbl.rows[0].cells[0], "Thủ trưởng đơn vị\n(Ký, họ tên)", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl.rows[0].cells[1], "Người đề nghị\n(Ký, họ tên)", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl.rows[1].cells[0], "\n\n{{lanh_dao}}", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl.rows[1].cells[1], "\n\n{{ho_ten}}", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.save(f"{OUT}/giay_de_nghi.docx")
    print("✓ giay_de_nghi.docx")

# ─────────────────────────────────────────────
# 3. BẢNG KÊ
# ─────────────────────────────────────────────
def make_bang_ke():
    doc = Document()
    set_margins(doc)

    # Header chung
    para(doc, "TỔNG CÔNG TY VIỄN THÔNG MOBIFONE", bold=True, size=12,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    para(doc, "Đơn vị: MOBIFONE LÂM ĐỒNG", size=12,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)

    para(doc, "BẢNG KÊ THANH TOÁN", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    para(doc, "CHI PHÍ NĂM {{nam_bc_so}}", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)

    para(doc, "Họ và tên:\t\t{{ho_ten}}", size=13, space_after=2)
    para(doc, "Đơn vị:\t\tMobiFone Lâm Đồng", size=13, space_after=8)

    # Bảng kê chi tiết
    headers = ["STT","Số chứng từ","Nội dung CV","Mã KM","Tên KM","Tài khoản","Nghiệp vụ","Mã SPDV",
               "Trước VAT","VAT","Sau VAT"]
    widths  = [Cm(0.8), Cm(1.5), Cm(3.5), Cm(1.5), Cm(4.5), Cm(1.5), Cm(1.2), Cm(1.0),
               Cm(2.0), Cm(1.5), Cm(2.0)]

    tbl = doc.add_table(rows=3, cols=len(headers))
    add_table_border(tbl)

    # Hàng header
    for i, (h, w) in enumerate(zip(headers, widths)):
        cell = tbl.rows[0].cells[i]
        cell.width = w
        cell_text(cell, h, bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Hàng dữ liệu
    row_data = ["1", "{{so_hd}}", "CHI TIEP KHACH T{{thang_tt}} - {{ma_kmcp}}",
                "{{ma_kmcp}}", "MOBIFONE - VT - Di động - CPQLDN - CP Tiếp khách, khánh tiết",
                "{{tk_kt}}", "{{nghiep_vu}}", "{{ma_spdv}}",
                "{{truoc_vat}}", "{{tien_vat}}", "{{sau_vat}}"]
    for i, val in enumerate(row_data):
        cell_text(tbl.rows[1].cells[i], val, size=10,
                  align=WD_ALIGN_PARAGRAPH.CENTER if i in [0,1,7,8,9,10] else WD_ALIGN_PARAGRAPH.LEFT)

    # Hàng tổng
    tbl.rows[2].cells[0].merge(tbl.rows[2].cells[7])
    cell_text(tbl.rows[2].cells[0], "Tổng cộng", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl.rows[2].cells[8], "{{truoc_vat}}", bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl.rows[2].cells[9], "{{tien_vat}}", bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl.rows[2].cells[10],"{{sau_vat}}",  bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)

    para(doc, "", space_after=4)
    para(doc, "\t\t\t\t\t{{ngay_bao_cao}}", italic=True, size=12,
         align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=4)

    tbl2 = doc.add_table(rows=2, cols=2)
    for row in tbl2.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcB = OxmlElement("w:tcBorders")
            for side in ["top","left","bottom","right"]:
                b = OxmlElement(f"w:{side}")
                b.set(qn("w:val"), "none")
                tcB.append(b)
            tcPr.append(tcB)

    cell_text(tbl2.rows[0].cells[0], "PHỤ TRÁCH CHI PHÍ ĐƠN VỊ\n(Ghi rõ họ, tên)", bold=True, size=12,
              align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl2.rows[0].cells[1], "NGƯỜI ĐỀ NGHỊ\n(Ghi rõ họ, tên)", bold=True, size=12,
              align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl2.rows[1].cells[0], "\n\n{{phu_trach_cp}}", bold=True, size=12,
              align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl2.rows[1].cells[1], "\n\n{{ho_ten}}", bold=True, size=12,
              align=WD_ALIGN_PARAGRAPH.CENTER)

    # Sheet 2: Bảng kê khai
    doc.add_page_break()
    para(doc, "TỔNG CÔNG TY VIỄN THÔNG MOBIFONE", bold=True, size=12,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    para(doc, "Đơn vị: MOBIFONE LÂM ĐỒNG", size=12,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    para(doc, "BẢNG KÊ KHAI CHI PHÍ THƯỜNG XUYÊN THANH TOÁN", bold=True, size=13,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    para(doc, "THÁNG {{thang_tk_so}} NĂM {{nam_tk_so}}", bold=True, size=13,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)

    hdrs2 = ["STT","Loại chi phí","Tháng thanh toán","Hóa đơn","Trước VAT","VAT","Sau VAT"]
    tbl3 = doc.add_table(rows=3, cols=len(hdrs2))
    add_table_border(tbl3)
    for i, h in enumerate(hdrs2):
        cell_text(tbl3.rows[0].cells[i], h, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    row2_data = ["1","Chi phí tiếp khách T{{thang_tk_so}}.{{nam_tk_so}}",
                 "{{thang_tt}}","{{so_hd}}","{{truoc_vat}}","{{tien_vat}}","{{sau_vat}}"]
    for i, val in enumerate(row2_data):
        cell_text(tbl3.rows[1].cells[i], val, size=11,
                  align=WD_ALIGN_PARAGRAPH.CENTER if i in [0,2,3,4,5,6] else WD_ALIGN_PARAGRAPH.LEFT)
    tbl3.rows[2].cells[0].merge(tbl3.rows[2].cells[3])
    cell_text(tbl3.rows[2].cells[0], "Tổng cộng", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl3.rows[2].cells[4], "{{truoc_vat}}", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl3.rows[2].cells[5], "{{tien_vat}}", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl3.rows[2].cells[6], "{{sau_vat}}", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)

    para(doc, "", space_after=6)
    para(doc, "\t\t\t\tNgười lập bảng", size=12, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=2)
    para(doc, "\n\n\t\t\t\t{{ho_ten}}", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.RIGHT)

    doc.save(f"{OUT}/bang_ke.docx")
    print("✓ bang_ke.docx")

# ─────────────────────────────────────────────
# 4. BÁO CÁO KẾT QUẢ CÔNG VIỆC
# ─────────────────────────────────────────────
def make_bao_cao_kqcv():
    doc = Document()
    set_margins(doc)

    two_col_header(doc,
        left_lines=[
            ("TỔNG CÔNG TY VIỄN THÔNG MOBIFONE", False, 11),
            ("MOBIFONE LÂM ĐỒNG", True, 11),
        ],
        right_lines=[
            ("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", True, 11),
            ("Độc lập – Tự do – Hạnh phúc", False, 11),
            ("───────────────────", False, 9),
            ("Lâm Đồng, {{ngay_bao_cao}}", False, 11),
        ]
    )

    para(doc, "BÁO CÁO KẾT QUẢ CÔNG VIỆC", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)

    para(doc, "Họ và tên: {{ho_ten}}", size=13, space_after=4)
    para(doc, "Tổ/ Bộ phận công tác: {{don_vi}}", size=13, space_after=4)
    para(doc, "Nội dung công việc: Tiếp khách", size=13, space_after=4)
    para(doc, "Thành phần: Lãnh đạo Chi nhánh, chuyên viên và khách mời (chi tiết theo tờ trình)",
         size=13, space_after=8)

    headers = ["STT","Nội dung công việc","Thời gian làm việc","Kết quả đạt được","Chưa đạt được","Hướng giải quyết"]
    widths  = [Cm(1.0), Cm(5.5), Cm(2.5), Cm(5.5), Cm(2.0), Cm(2.5)]

    tbl = doc.add_table(rows=2, cols=len(headers))
    add_table_border(tbl)
    for i, (h, w) in enumerate(zip(headers, widths)):
        c = tbl.rows[0].cells[i]
        c.width = w
        cell_text(c, h, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)

    row_data = [
        "01",
        "{{ly_do}}",
        "{{ngay_tk_so}}/{{thang_tk_so}}/{{nam_tk_so}}",
        "{{ket_qua}}",
        "",
        ""
    ]
    for i, val in enumerate(row_data):
        cell_text(tbl.rows[1].cells[i], val, size=11,
                  align=WD_ALIGN_PARAGRAPH.CENTER if i in [0,2] else WD_ALIGN_PARAGRAPH.LEFT)

    para(doc, "", space_after=10)

    tbl2 = doc.add_table(rows=2, cols=2)
    for row in tbl2.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcB = OxmlElement("w:tcBorders")
            for side in ["top","left","bottom","right"]:
                b = OxmlElement(f"w:{side}")
                b.set(qn("w:val"), "none")
                tcB.append(b)
            tcPr.append(tcB)

    cell_text(tbl2.rows[0].cells[0], "", size=12)
    cell_text(tbl2.rows[0].cells[1], "NGƯỜI BÁO CÁO", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(tbl2.rows[1].cells[0], "", size=12)
    cell_text(tbl2.rows[1].cells[1], "\n\n{{ho_ten}}", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.save(f"{OUT}/bao_cao_kqcv.docx")
    print("✓ bao_cao_kqcv.docx")

if __name__ == "__main__":
    make_to_trinh()
    make_giay_de_nghi()
    make_bang_ke()
    make_bao_cao_kqcv()
    print("\nĐã tạo xong 4 file template trong thư mục word_templates/")
