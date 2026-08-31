#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script tạo file Word hướng dẫn sử dụng mynhc (ClassReg) — phiên bản rút gọn
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import docx.enum.text

doc = Document()

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def set_page_margins(doc, top=2.0, bottom=2.0, left=2.5, right=2.0):
    for section in doc.sections:
        section.top_margin    = Cm(top)
        section.bottom_margin = Cm(bottom)
        section.left_margin   = Cm(left)
        section.right_margin  = Cm(right)

def set_font(run, name="Times New Roman", size=12, bold=False, italic=False, color=None):
    run.font.name  = name
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)
    r = run._r
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rFonts.set(qn('w:cs'), name)
    existing = rPr.find(qn('w:rFonts'))
    if existing is not None:
        rPr.remove(existing)
    rPr.insert(0, rFonts)

def heading(doc, text, level=1, center=False):
    style_map = {1: 'Heading 1', 2: 'Heading 2'}
    p = doc.add_paragraph(style=style_map.get(level, 'Heading 1'))
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.bold = True
    if level == 1:
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    else:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x0E, 0x74, 0xB5)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = run._r
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rFonts.set(qn(attr), "Times New Roman")
    existing = rPr.find(qn('w:rFonts'))
    if existing is not None:
        rPr.remove(existing)
    rPr.insert(0, rFonts)
    return p

def body(doc, text, bold=False, italic=False, indent=False, color=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_font(run, bold=bold, italic=italic, color=color)
    if indent:
        p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.space_after = Pt(3)
    return p

def bullet(doc, text, indent=0.8):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    set_font(run)

def note(doc, text, title="Lưu ý", bg="FFF2CC", tc=(0xBF,0x8F,0x00)):
    def shade(p, hex_c):
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_c)
        pPr.append(shd)
    pt = doc.add_paragraph()
    pt.paragraph_format.left_indent  = Cm(0.4)
    pt.paragraph_format.right_indent = Cm(0.4)
    pt.paragraph_format.space_before = Pt(4)
    pt.paragraph_format.space_after  = Pt(0)
    set_font(pt.add_run(f"  {title}"), bold=True, size=11, color=tc)
    shade(pt, bg)
    pb = doc.add_paragraph()
    pb.paragraph_format.left_indent  = Cm(0.4)
    pb.paragraph_format.right_indent = Cm(0.4)
    pb.paragraph_format.space_before = Pt(0)
    pb.paragraph_format.space_after  = Pt(5)
    set_font(pb.add_run(f"  {text}"), size=11, italic=True)
    shade(pb, bg)

def table(doc, headers, rows, widths=None):
    t = doc.add_table(rows=1+len(rows), cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hrow = t.rows[0]
    for i, h in enumerate(headers):
        cell = hrow.cells[i]
        cell.text = h
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'1E4FA3')
        tcPr.append(shd)
        for run in cell.paragraphs[0].runs:
            set_font(run, bold=True, size=11, color=(0xFF,0xFF,0xFF))
    for ri, row_data in enumerate(rows):
        row = t.rows[ri+1]
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            cell.text = str(val)
            for run in cell.paragraphs[0].runs:
                set_font(run, size=11)
            if ri % 2 == 1:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'EEF4FB')
                tcPr.append(shd)
    if widths:
        for row in t.rows:
            for i, cell in enumerate(row.cells):
                cell.width = Cm(widths[i])
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t

def add_header_footer(doc):
    for section in doc.sections:
        header = section.header
        header.is_linked_to_previous = False
        hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        hp.clear(); hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(hp.add_run("Hướng dẫn sử dụng hệ thống mynhc — Trường THPT Nguyễn Hữu Cầu"), size=10, italic=True, color=(0x44,0x72,0xC4))
        pPr = hp._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bot = OxmlElement('w:bottom')
        bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'4'); bot.set(qn('w:space'),'1'); bot.set(qn('w:color'),'4472C4')
        pBdr.append(bot); pPr.append(pBdr)
        footer = section.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        fp.clear(); fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(fp.add_run("Trang "), size=10)
        fc1 = OxmlElement('w:fldChar'); fc1.set(qn('w:fldCharType'), 'begin')
        instr = OxmlElement('w:instrText'); instr.text = ' PAGE '
        fc2 = OxmlElement('w:fldChar'); fc2.set(qn('w:fldCharType'), 'separate')
        fc3 = OxmlElement('w:fldChar'); fc3.set(qn('w:fldCharType'), 'end')
        r_pg = fp.add_run()
        r_pg._r.append(fc1); r_pg._r.append(instr)
        r_pg._r.append(fc2); r_pg._r.append(fc3)
        set_font(r_pg, size=10)

def page_break(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(docx.enum.text.WD_BREAK.PAGE)

# ─────────────────────────────────────────────────────────────────────────────
# PAGE SETUP
# ─────────────────────────────────────────────────────────────────────────────
set_page_margins(doc)

# ─────────────────────────────────────────────────────────────────────────────
# BÌA
# ─────────────────────────────────────────────────────────────────────────────
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HỒ CHÍ MINH"), bold=True, size=13)

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("TRƯỜNG THPT NGUYỄN HỮU CẦU"), bold=True, size=14, color=(0x1F,0x49,0x7D))

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("─" * 38), size=11)

for _ in range(3): doc.add_paragraph()

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("HỆ THỐNG ĐĂNG KÝ LỚP HỌC"), bold=True, size=20, color=(0x1F,0x49,0x7D))

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("mynhc"), bold=True, size=28, color=(0x2E,0x74,0xB5))

doc.add_paragraph()

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("HƯỚNG DẪN SỬ DỤNG NHANH"), bold=True, size=16, color=(0xBF,0x8F,0x00))

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("Dành cho Quản trị viên – Giáo viên – Học sinh"), italic=True, size=12, color=(0x40,0x40,0x40))

for _ in range(5): doc.add_paragraph()

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("Năm học 2026 – 2027"), bold=True, size=13)

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("Phiên bản 1.1  |  Tháng 8 năm 2026"), italic=True, size=11, color=(0x60,0x60,0x60))

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# 1. TỔNG QUAN
# ─────────────────────────────────────────────────────────────────────────────
heading(doc, "1. TỔNG QUAN")
body(doc, "Hệ thống mynhc hỗ trợ đăng ký lớp học tự chọn theo hai giai đoạn:")
bullet(doc, "Giai đoạn 1: Giáo viên đăng ký mở lớp → Admin điền phòng & sĩ số.")
bullet(doc, "Giai đoạn 2: Học sinh đăng ký môn học theo khối, hệ thống kiểm tra trùng lịch tự động.")

table(doc,
    ["Vai trò", "URL đăng nhập", "Chức năng chính"],
    [
        ["Quản trị viên", "/admin", "Quản lý GV, HS, điều phối giai đoạn, phân phòng, xuất Excel"],
        ["Giáo viên",     "/login (chọn GV)", "Đăng ký lịch mở lớp, xem số HS & phòng cập nhật trực tiếp"],
        ["Học sinh",      "/login (chọn HS)", "Đăng ký / huỷ môn học, xem thời khoá biểu cá nhân"],
    ],
    widths=[3.5, 4.5, 8.5]
)

note(doc,
    "Lớp chỉ xuất hiện cho học sinh khi đã được điền đầy đủ Địa điểm VÀ Sĩ số tối đa.",
    title="Điều kiện hiển thị lớp",
    bg="DEEAF1", tc=(0x1F,0x49,0x7D)
)

# ─────────────────────────────────────────────────────────────────────────────
# 2. ĐĂNG NHẬP
# ─────────────────────────────────────────────────────────────────────────────
heading(doc, "2. ĐĂNG NHẬP")
table(doc,
    ["Vai trò", "Thông tin đăng nhập", "Mật khẩu mặc định (kiểm thử)"],
    [
        ["Quản trị viên", "Chọn Quản trị viên + nhập mật khẩu", "Admin@123"],
        ["Giáo viên",     "Chọn Giáo viên + nhập CCCD",          "Nhc@2627"],
        ["Học sinh",      "Chọn Học sinh + nhập CCCD",            "Nhc@2627"],
    ],
    widths=[3.5, 7.5, 5.5]
)
note(doc,
    "Tài khoản mới tạo (chưa kích hoạt) phải đăng nhập lần đầu để đặt mật khẩu. "
    "Khi Admin khôi phục MK, user nhận mật khẩu tạm thời 12 ký tự ngẫu nhiên và bắt buộc đổi ngay sau khi đăng nhập.",
    title="Kích hoạt tài khoản"
)
body(doc, "Yêu cầu mật khẩu: tối thiểu 8 ký tự, gồm chữ hoa/thường/số/ký tự đặc biệt (!@#$%^&*...). "
          "Không dùng: dấu nháy đơn ('), nháy kép (\"), chấm phẩy (;), gạch chéo (\\).",
     italic=True, color=(0x47,0x55,0x69))

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# 3. QUẢN TRỊ VIÊN
# ─────────────────────────────────────────────────────────────────────────────
heading(doc, "3. QUẢN TRỊ VIÊN (/admin)")

heading(doc, "3.1. Quản lý Giáo viên & Học sinh", 2)
table(doc,
    ["Tính năng", "Mô tả"],
    [
        ["Xem danh sách", "Bảng hiển thị CCCD, Họ tên, Email, Trạng thái TK (Đã / Chưa kích hoạt)"],
        ["Tìm kiếm", "Ô search lọc theo Họ tên hoặc CCCD ngay lập tức"],
        ["Thêm / Sửa / Xoá", "Nút icon trong cột Thao tác (chỉ hiện với tài khoản đã kích hoạt)"],
        ["Nhập Excel", "Upload file .xlsx theo mẫu; hỗ trợ cả GV lẫn HS"],
        ["Khôi phục MK", "Tạo mật khẩu tạm 12 ký tự ngẫu nhiên, hiển thị 1 lần — sao chép trước khi đóng modal"],
    ],
    widths=[4.5, 12.0]
)

heading(doc, "3.2. Điều phối Giai đoạn", 2)
table(doc,
    ["Giai đoạn", "Ai làm gì", "Nút điều khiển"],
    [
        ["Giai đoạn 1 – MỞ",   "Giáo viên đăng ký lịch mở lớp",                          "Mở đăng ký GV  →  /admin/class-reg"],
        ["Giai đoạn 1 – ĐÓNG", "Admin điền phòng & sĩ số tối đa cho từng lớp",            "Đóng đăng ký GV"],
        ["Giai đoạn 2 – MỞ",   "Học sinh đăng ký / huỷ môn học",                          "Mở đăng ký HS  →  /admin/enrollment"],
        ["Giai đoạn 2 – ĐÓNG", "Xuất Excel kết quả; HS chỉ xem TKB, không thể thay đổi", "Đóng đăng ký HS"],
    ],
    widths=[4.0, 7.5, 5.0]
)
note(doc,
    "Lớp có viền đỏ = chưa có phòng hoặc sĩ số tối đa → cần bổ sung trước khi mở Giai đoạn 2. "
    "Phòng & sĩ số lưu tự động khi nhấn Enter hoặc click ra ngoài ô.",
    title="Lưu ý Giai đoạn 1"
)

heading(doc, "3.3. Quản lý Đăng ký Môn học (/admin/enrollment)", 2)
body(doc, "Bảng hiển thị tất cả lớp đã publish: ID, GV, Môn, Khối, "
          "Thời gian (Thứ X – Sáng/Chiều – Tiết X–Y), Địa điểm, Sĩ số (cập nhật mỗi 5 giây).")
body(doc, "Nhấn Xem HS → modal danh sách học sinh đã đăng ký (STT, Họ tên, CCCD, Khối, Lớp).")
body(doc, "Thanh công cụ phía trên bảng: Thêm lớp thủ công · Xuất Excel · Upload phân phòng.")

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# 4. GIÁO VIÊN
# ─────────────────────────────────────────────────────────────────────────────
heading(doc, "4. GIÁO VIÊN (/teacher/schedule)")
table(doc,
    ["Trạng thái", "Giáo viên có thể làm gì"],
    [
        ["Giai đoạn 1 đang MỞ",   "Click ô lịch trống → chọn Khối & số tiết → đăng ký lớp. Hover ô đã có → nút X để xoá."],
        ["Giai đoạn 1 đã ĐÓNG",   "Chỉ xem thời khoá biểu, không thể thêm/xoá lớp."],
        ["Giai đoạn 2 đang MỞ",   "Xem TKB; badge số HS/phòng cập nhật trực tiếp mỗi 5 giây."],
        ["Giai đoạn 2 đã ĐÓNG",   "Chỉ xem TKB và số HS cuối cùng."],
    ],
    widths=[5.0, 11.5]
)
note(doc, "Một ô lịch chỉ có thể đăng ký một lớp. Xoá lớp sẽ xoá toàn bộ đăng ký HS trong lớp đó.", title="Lưu ý")

# ─────────────────────────────────────────────────────────────────────────────
# 5. HỌC SINH
# ─────────────────────────────────────────────────────────────────────────────
heading(doc, "5. HỌC SINH (/student/schedule)")
table(doc,
    ["Trạng thái", "Học sinh có thể làm gì"],
    [
        ["Giai đoạn 2 chưa mở",  "Chỉ xem Tab Thời khoá biểu. Nút Đăng ký bị ẩn."],
        ["Giai đoạn 2 đang MỞ",  "Đăng ký và huỷ môn học. Badge sĩ số cập nhật tự động. Hệ thống kiểm tra trùng lịch."],
        ["Giai đoạn 2 đã ĐÓNG",  "Chỉ xem TKB cá nhân — không đăng ký, không huỷ, không tương tác."],
    ],
    widths=[5.0, 11.5]
)
body(doc, "Badge sĩ số: Xanh = còn chỗ · Vàng = gần đầy (≥ 2/3) · Xám = đã đầy (không đăng ký được).", italic=True)
note(doc,
    "Lần đầu đăng nhập hoặc sau khi Admin khôi phục MK: hệ thống tự chuyển đến trang "
    "Đổi mật khẩu bắt buộc trước khi vào trang chính.",
    title="Đổi mật khẩu bắt buộc"
)

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# 6. XỬ LÝ SỰ CỐ THƯỜNG GẶP
# ─────────────────────────────────────────────────────────────────────────────
heading(doc, "6. XỬ LÝ SỰ CỐ THƯỜNG GẶP")
table(doc,
    ["Sự cố", "Nguyên nhân", "Cách xử lý"],
    [
        ["Sai mật khẩu",
         "Nhập sai hoặc quên MK",
         "Liên hệ Admin khôi phục MK → nhận MK tạm → đổi ngay"],
        ["GV không click được ô lịch",
         "Giai đoạn 1 chưa mở hoặc đã đóng",
         "Admin vào /admin/class-reg → Mở đăng ký GV"],
        ["HS không thấy nút Đăng ký",
         "Giai đoạn 2 chưa mở hoặc lớp đã đầy (badge xám)",
         "Kiểm tra badge màu; nếu không xám → Admin mở Giai đoạn 2"],
        ["Lớp không hiện với HS",
         "Thiếu địa điểm hoặc sĩ số tối đa",
         "Admin vào /admin/class-reg → điền phòng & sĩ số"],
        ["Lỗi nhập Excel",
         "Sai định dạng cột hoặc dữ liệu trùng",
         "Tải file mẫu mới; không đổi tên cột; kiểm tra CCCD không trùng"],
        ["Trang tải chậm / lỗi kết nối",
         "Sự cố mạng hoặc máy chủ",
         "Nhấn F5 tải lại; báo bộ phận kỹ thuật nếu kéo dài"],
    ],
    widths=[4.5, 5.0, 7.0]
)

doc.add_paragraph()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("─" * 45), size=10, color=(0x80,0x80,0x80))

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("Tài liệu thuộc Trường THPT Nguyễn Hữu Cầu  ·  mynhc v1.1  ·  Năm học 2026–2027"),
         italic=True, size=10, color=(0x60,0x60,0x60))

# ─────────────────────────────────────────────────────────────────────────────
add_header_footer(doc)
# ─────────────────────────────────────────────────────────────────────────────

output_path = "/Users/ddang/Documents/Subject_Survey/ClassReg/Hướng_dẫn_sử_dụng_mynhc.docx"
doc.save(output_path)
print(f"Đã lưu: {output_path}")
