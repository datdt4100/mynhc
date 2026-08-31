#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script tạo file Word hướng dẫn sử dụng mynhc (ClassReg)
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
    run.font.name   = name
    run.font.size   = Pt(size)
    run.font.bold   = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)
    r    = run._r
    rPr  = r.get_or_add_rPr()
    rFnt = OxmlElement('w:rFonts')
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rFnt.set(qn(attr), name)
    existing = rPr.find(qn('w:rFonts'))
    if existing is not None:
        rPr.remove(existing)
    rPr.insert(0, rFnt)

def H1(doc, text):
    p   = doc.add_paragraph(style='Heading 1')
    run = p.add_run(text)
    run.font.name  = "Times New Roman"; run.font.bold = True
    run.font.size  = Pt(14)
    run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    _fix_font(run)
    return p

def H2(doc, text):
    p   = doc.add_paragraph(style='Heading 2')
    run = p.add_run(text)
    run.font.name  = "Times New Roman"; run.font.bold = True
    run.font.size  = Pt(12)
    run.font.color.rgb = RGBColor(0x0E, 0x74, 0xB5)
    _fix_font(run)
    return p

def H3(doc, text):
    p   = doc.add_paragraph(style='Heading 3')
    run = p.add_run(text)
    run.font.name  = "Times New Roman"; run.font.bold = True
    run.font.size  = Pt(12)
    run.font.color.rgb = RGBColor(0x1F, 0x74, 0x89)
    _fix_font(run)
    return p

def _fix_font(run):
    r   = run._r
    rPr = r.get_or_add_rPr()
    rFnt = OxmlElement('w:rFonts')
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rFnt.set(qn(attr), "Times New Roman")
    existing = rPr.find(qn('w:rFonts'))
    if existing is not None:
        rPr.remove(existing)
    rPr.insert(0, rFnt)

def body(doc, text, bold=False, italic=False, indent=False, color=None):
    p   = doc.add_paragraph()
    run = p.add_run(text)
    set_font(run, bold=bold, italic=italic, color=color)
    if indent:
        p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.space_after = Pt(3)
    return p

def step(doc, number, text):
    p = doc.add_paragraph(style='List Number')
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.space_after = Pt(2)
    set_font(p.add_run(text))

def bullet(doc, text, indent=0.8):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(2)
    set_font(p.add_run(text))

def note(doc, text, title="Lưu ý", bg="FFF2CC", tc=(0xBF, 0x8F, 0x00)):
    def shade(p, hex_c):
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), hex_c)
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
    pb.paragraph_format.space_after  = Pt(6)
    set_font(pb.add_run(f"  {text}"), size=11, italic=True)
    shade(pb, bg)

def tbl(doc, headers, rows, widths=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hrow = t.rows[0]
    for i, h in enumerate(headers):
        cell = hrow.cells[i]
        cell.text = h
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        tcEl  = cell._tc
        tcPr  = tcEl.get_or_add_tcPr()
        shd   = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '1E4FA3')
        tcPr.append(shd)
        for run in cell.paragraphs[0].runs:
            set_font(run, bold=True, size=11, color=(0xFF, 0xFF, 0xFF))
    for ri, row_data in enumerate(rows):
        row = t.rows[ri + 1]
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            cell.text = str(val)
            for run in cell.paragraphs[0].runs:
                set_font(run, size=11)
            if ri % 2 == 1:
                tcEl = cell._tc
                tcPr = tcEl.get_or_add_tcPr()
                shd  = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), 'EEF4FB')
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
        hp.clear()
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(hp.add_run("Hướng dẫn sử dụng hệ thống mynhc — Trường THPT Nguyễn Hữu Cầu"),
                 size=10, italic=True, color=(0x44, 0x72, 0xC4))
        pPr = hp._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bot  = OxmlElement('w:bottom')
        bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), '4')
        bot.set(qn('w:space'), '1');    bot.set(qn('w:color'), '4472C4')
        pBdr.append(bot); pPr.append(pBdr)

        footer = section.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        fp.clear()
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(fp.add_run("Trang "), size=10)
        fc1   = OxmlElement('w:fldChar'); fc1.set(qn('w:fldCharType'), 'begin')
        instr = OxmlElement('w:instrText'); instr.text = ' PAGE '
        fc2   = OxmlElement('w:fldChar'); fc2.set(qn('w:fldCharType'), 'separate')
        fc3   = OxmlElement('w:fldChar'); fc3.set(qn('w:fldCharType'), 'end')
        r_pg  = fp.add_run()
        r_pg._r.append(fc1); r_pg._r.append(instr)
        r_pg._r.append(fc2); r_pg._r.append(fc3)
        set_font(r_pg, size=10)

def page_break(doc):
    doc.add_paragraph().add_run().add_break(docx.enum.text.WD_BREAK.PAGE)

# ─────────────────────────────────────────────────────────────────────────────
set_page_margins(doc)

# ─────────────────────────────────────────────────────────────────────────────
# BÌA
# ─────────────────────────────────────────────────────────────────────────────
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HỒ CHÍ MINH"), bold=True, size=13)

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("TRƯỜNG THPT NGUYỄN HỮU CẦU"), bold=True, size=14, color=(0x1F, 0x49, 0x7D))

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("─" * 38), size=11)

for _ in range(3):
    doc.add_paragraph()

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("HỆ THỐNG ĐĂNG KÝ LỚP HỌC"), bold=True, size=22, color=(0x1F, 0x49, 0x7D))

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("mynhc"), bold=True, size=30, color=(0x2E, 0x74, 0xB5))

doc.add_paragraph()

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("HƯỚNG DẪN SỬ DỤNG"), bold=True, size=18, color=(0xBF, 0x8F, 0x00))

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("Dành cho Quản trị viên – Giáo viên – Học sinh"), italic=True, size=13, color=(0x40, 0x40, 0x40))

for _ in range(5):
    doc.add_paragraph()

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("Năm học 2026 – 2027"), bold=True, size=14)

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("Phiên bản 1.1  |  Tháng 8 năm 2026"), italic=True, size=11, color=(0x60, 0x60, 0x60))

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# 1. TỔNG QUAN
# ─────────────────────────────────────────────────────────────────────────────
H1(doc, "1. TỔNG QUAN HỆ THỐNG")
body(doc, "Hệ thống mynhc là nền tảng đăng ký lớp học tự chọn trực tuyến dành riêng cho Trường THPT Nguyễn Hữu Cầu. Quy trình vận hành gồm hai giai đoạn:")
bullet(doc, "Giai đoạn 1: Giáo viên đăng ký lịch mở lớp → Quản trị viên điền phòng học & sĩ số tối đa.")
bullet(doc, "Giai đoạn 2: Học sinh đăng ký môn học theo khối; hệ thống kiểm tra trùng lịch tự động.")

tbl(doc,
    ["Vai trò", "URL truy cập", "Chức năng chính"],
    [
        ["Quản trị viên", "/admin",           "Quản lý GV, HS; điều phối giai đoạn; phân phòng; xuất Excel"],
        ["Giáo viên",     "/login → chọn GV", "Đăng ký lịch mở lớp; xem số HS & phòng học theo thời gian thực"],
        ["Học sinh",      "/login → chọn HS", "Xem thời khoá biểu; đăng ký và huỷ môn học"],
    ],
    widths=[3.5, 4.0, 9.0]
)

note(doc,
    "Lớp học chỉ hiển thị cho học sinh khi đã được điền đầy đủ Địa điểm VÀ Sĩ số tối đa.",
    title="Quan trọng", bg="FCE4D6", tc=(0xC0, 0x50, 0x00)
)

tbl(doc,
    ["Vai trò", "Thông tin đăng nhập", "Mật khẩu mặc định (kiểm thử)"],
    [
        ["Quản trị viên", "Chọn Quản trị viên + nhập mật khẩu", "Admin@123"],
        ["Giáo viên",     "Chọn Giáo viên + nhập CCCD",          "Nhc@2627"],
        ["Học sinh",      "Chọn Học sinh + nhập CCCD",            "Nhc@2627"],
    ],
    widths=[3.5, 7.5, 5.5]
)

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# 2. ĐĂNG NHẬP
# ─────────────────────────────────────────────────────────────────────────────
H1(doc, "2. ĐĂNG NHẬP VÀ MẬT KHẨU")

H2(doc, "2.1. Đăng nhập")
body(doc, "Tất cả người dùng đăng nhập tại cùng một trang /login:")
step(doc, 1, "Truy cập URL hệ thống (ví dụ: https://mynhc.onrender.com/login).")
step(doc, 2, "Chọn loại người dùng trong dropdown: Quản trị viên / Giáo viên / Học sinh.")
step(doc, 3, "Nhập CCCD (GV/HS) hoặc mật khẩu (Admin).")
step(doc, 4, "Nhấn Đăng nhập. Hệ thống chuyển đến trang tương ứng.")

note(doc,
    "Nếu tài khoản chưa kích hoạt (lần đầu đăng nhập hoặc vừa được Admin khôi phục MK), "
    "hệ thống tự chuyển đến trang Đổi mật khẩu bắt buộc trước khi vào trang chính.",
    title="Kích hoạt tài khoản"
)

H2(doc, "2.2. Đặt / Đổi mật khẩu bắt buộc")
body(doc, "Xảy ra khi: (a) đăng nhập lần đầu, hoặc (b) Admin vừa khôi phục mật khẩu.")
step(doc, 1, "Đọc hộp thông tin xanh hiển thị yêu cầu mật khẩu.")
step(doc, 2, "Nhập mật khẩu mới vào ô Mật khẩu mới.")
step(doc, 3, "Nhập lại vào ô Xác nhận mật khẩu (ô hiển thị ✓ Trùng khớp khi đúng).")
step(doc, 4, "Nhấn Xác nhận đổi mật khẩu. Hệ thống chuyển về trang chính.")

note(doc,
    "Yêu cầu mật khẩu: tối thiểu 8 ký tự; gồm chữ hoa (A–Z), chữ thường (a–z), "
    "chữ số (0–9) và ký tự đặc biệt (!@#$%^&*()-_+=[]{} |<>,.?/~). "
    "KHÔNG dùng: nháy đơn ('), nháy kép (\"), chấm phẩy (;), gạch chéo ngược (\\).",
    title="Yêu cầu mật khẩu", bg="DEEAF1", tc=(0x1F, 0x49, 0x7D)
)

H2(doc, "2.3. Đăng xuất")
step(doc, 1, "Nhấn nút Đăng xuất ở góc phải trên cùng của trang.")
step(doc, 2, "Hệ thống xoá phiên làm việc và chuyển về trang đăng nhập.")

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# 3. QUẢN TRỊ VIÊN
# ─────────────────────────────────────────────────────────────────────────────
H1(doc, "3. HƯỚNG DẪN QUẢN TRỊ VIÊN")
body(doc, "Sau khi đăng nhập, quản trị viên được chuyển đến trang /admin. Trang chủ hiển thị 4 thẻ thống kê (số GV, HS, lớp đã mở, lượt đăng ký) và trạng thái hai giai đoạn.")

H2(doc, "3.1. Quản lý Giáo viên")

H3(doc, "3.1.1. Xem danh sách")
body(doc, "Bảng danh sách gồm: STT, CCCD, Họ tên, Giới tính, Tổ bộ môn, Email, Trạng thái TK (Đã/Chưa kích hoạt), Thao tác.")
body(doc, "Ô tìm kiếm phía trên bảng lọc theo Họ tên hoặc CCCD ngay lập tức.")

H3(doc, "3.1.2. Thêm giáo viên")
step(doc, 1, "Nhấn nút Thêm GV phía trên bảng.")
step(doc, 2, "Điền thông tin: CCCD, Họ tên, Giới tính, Tổ bộ môn, Email (tuỳ chọn), Mật khẩu.")
step(doc, 3, "Nhấn Lưu. Thông báo xác nhận xuất hiện.")

H3(doc, "3.1.3. Sửa thông tin")
step(doc, 1, "Nhấn nút icon Sửa (bút chì) ở cột Thao tác.")
step(doc, 2, "Cập nhật thông tin trong form.")
step(doc, 3, "Nhấn Lưu thay đổi.")

H3(doc, "3.1.4. Xoá giáo viên")
step(doc, 1, "Nhấn nút icon Xoá (thùng rác) ở cột Thao tác.")
step(doc, 2, "Xác nhận trong hộp thoại. Xoá GV sẽ xoá toàn bộ lớp học GV đó đã đăng ký.")

H3(doc, "3.1.5. Khôi phục mật khẩu")
body(doc, "Chỉ hiện với tài khoản đã kích hoạt:")
step(doc, 1, "Nhấn nút icon chìa khoá (màu cam) ở cột Thao tác.")
step(doc, 2, "Xác nhận trong hộp thoại.")
step(doc, 3, "Hệ thống tạo mật khẩu tạm thời 12 ký tự ngẫu nhiên, hiển thị trong modal màu vàng.")
step(doc, 4, "Sao chép mật khẩu (nút Copy) và cung cấp cho giáo viên.")
step(doc, 5, "Giáo viên đăng nhập bằng MK tạm → hệ thống tự chuyển đến trang Đổi MK bắt buộc.")
note(doc, "Mật khẩu tạm CHỈ hiển thị một lần. Sao chép ngay trước khi đóng modal.",
     title="Quan trọng", bg="FCE4D6", tc=(0xC0, 0x50, 0x00))

H3(doc, "3.1.6. Nhập danh sách từ Excel")
step(doc, 1, "Nhấn nút Tải mẫu Excel để tải file mẫu về máy.")
step(doc, 2, "Điền dữ liệu vào file mẫu (CCCD, Họ tên, Giới tính, Tổ bộ môn, Email, MK).")
step(doc, 3, "Nhấn Nhập Excel, chọn file đã điền.")
step(doc, 4, "Hệ thống xử lý và hiển thị báo cáo: số bản ghi thành công / lỗi và lý do.")

H2(doc, "3.2. Quản lý Học sinh")
body(doc, "Thao tác tương tự Quản lý Giáo viên. Bảng gồm: STT, CCCD, Họ tên, Khối, Lớp, Email, Trạng thái TK, Thao tác.")
body(doc, "File Excel mẫu gồm các cột: CCCD, Họ tên, Khối (10/11/12), Lớp, Email, Mật khẩu.")

page_break(doc)

H2(doc, "3.3. Quản lý Đăng ký Mở lớp — Giai đoạn 1 (/admin/class-reg)")

H3(doc, "3.3.1. Mở và đóng Giai đoạn 1")
step(doc, 1, "Vào Admin → Quản lý đăng ký mở lớp (hoặc URL /admin/class-reg).")
step(doc, 2, "Nhấn Mở đăng ký GV để giáo viên bắt đầu đăng ký lịch.")
step(doc, 3, "Sau khi thu thập đủ đăng ký, nhấn Đóng đăng ký GV.")
note(doc, "Sau khi đóng, giáo viên không thể thêm hoặc xoá lớp.", title="Lưu ý")

H3(doc, "3.3.2. Điền phòng học và sĩ số tối đa")
step(doc, 1, "Nhấn vào tên giáo viên để xem danh sách lớp đã đăng ký.")
step(doc, 2, "Tại cột Địa điểm, nhập tên phòng học (ví dụ: A201) rồi nhấn Enter hoặc click ra ngoài.")
step(doc, 3, "Tại cột Sĩ số, nhập số học sinh tối đa rồi nhấn Enter hoặc click ra ngoài.")
step(doc, 4, "Dữ liệu lưu tự động — ô chuyển màu xanh nhạt khi lưu thành công.")
step(doc, 5, "Lớp có viền đỏ = còn thiếu thông tin → cần bổ sung trước khi mở Giai đoạn 2.")

H3(doc, "3.3.3. Xuất / Upload file phân phòng")
bullet(doc, "Xuất Excel: nhấn Tải Excel phân phòng → file .xlsx tải về máy chứa toàn bộ thông tin lớp.")
bullet(doc, "Upload: điền phòng/sĩ số vào file Excel → nhấn Upload phân phòng → hệ thống cập nhật hàng loạt.")

H2(doc, "3.4. Quản lý Đăng ký Môn học — Giai đoạn 2 (/admin/enrollment)")

H3(doc, "3.4.1. Mở và đóng Giai đoạn 2")
step(doc, 1, "Kiểm tra: tất cả lớp phải có Địa điểm và Sĩ số tối đa (không còn viền đỏ).")
step(doc, 2, "Nhấn Mở đăng ký HS để học sinh bắt đầu đăng ký.")
step(doc, 3, "Sau thời hạn, nhấn Đóng đăng ký HS.")
note(doc, "Khi Giai đoạn 2 mở, học sinh có thể đăng ký VÀ huỷ. Sau khi đóng, không ai thay đổi được.",
     title="Lưu ý quan trọng", bg="FCE4D6", tc=(0xC0, 0x50, 0x00))

H3(doc, "3.4.2. Xem danh sách học sinh theo lớp")
body(doc, "Bảng hiển thị: ID lớp, Giáo viên, Môn/Tổ, Khối, Thời gian (Thứ X – Sáng/Chiều – Tiết X–Y), Địa điểm, Sĩ số (cập nhật tự động mỗi 5 giây).")
step(doc, 1, "Nhấn nút Xem HS ở cột Thao tác.")
step(doc, 2, "Modal hiển thị danh sách HS đã đăng ký: STT, Họ tên, CCCD, Khối, Lớp.")

H3(doc, "3.4.3. Xuất kết quả đăng ký")
step(doc, 1, "Nhấn Xuất Excel ở thanh công cụ phía trên bảng.")
step(doc, 2, "File .xlsx tải về chứa danh sách từng lớp và HS đã đăng ký.")

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# 4. GIÁO VIÊN
# ─────────────────────────────────────────────────────────────────────────────
H1(doc, "4. HƯỚNG DẪN GIÁO VIÊN (/teacher/schedule)")
body(doc, "Sau khi đăng nhập, giáo viên thấy bảng thời khoá biểu dạng lưới (Thứ 2–7, Buổi Sáng/Chiều, Tiết 1–4). Ô trống = chưa có lớp; ô tô màu = đã đăng ký lớp.")

tbl(doc,
    ["Trạng thái hệ thống", "Giáo viên có thể làm gì"],
    [
        ["Giai đoạn 1 đang MỞ",  "Click ô lịch trống để đăng ký lớp. Hover ô đã có → nút X để xoá."],
        ["Giai đoạn 1 đã ĐÓNG",  "Chỉ xem thời khoá biểu — không thể thêm hoặc xoá lớp."],
        ["Giai đoạn 2 đang MỞ",  "Xem TKB; badge số HS và phòng học cập nhật trực tiếp mỗi 5 giây."],
        ["Giai đoạn 2 đã ĐÓNG",  "Chỉ xem TKB và số HS cuối cùng."],
    ],
    widths=[5.0, 11.5]
)

H2(doc, "4.1. Đăng ký mở lớp")
note(doc, "Chỉ thực hiện được khi Giai đoạn 1 đang MỞ.", title="Điều kiện", bg="DEEAF1", tc=(0x1F, 0x49, 0x7D))
step(doc, 1, "Click vào ô lịch trống tương ứng (thứ, buổi, tiết muốn dạy).")
step(doc, 2, "Modal đăng ký hiện ra. Chọn Khối (10 / 11 / 12) và nhập Số tiết.")
step(doc, 3, "Nhấn Xác nhận. Ô lịch chuyển màu và hiển thị thông tin lớp vừa đăng ký.")

H2(doc, "4.2. Xoá lớp đã đăng ký")
note(doc, "Chỉ thực hiện được khi Giai đoạn 1 đang MỞ.", title="Điều kiện", bg="DEEAF1", tc=(0x1F, 0x49, 0x7D))
step(doc, 1, "Di chuột (hover) vào ô lịch đã có lớp.")
step(doc, 2, "Nhấn nút X màu đỏ xuất hiện ở góc trên phải của ô.")
step(doc, 3, "Xác nhận trong hộp thoại. Ô trở về trạng thái trống.")

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# 5. HỌC SINH
# ─────────────────────────────────────────────────────────────────────────────
H1(doc, "5. HƯỚNG DẪN HỌC SINH (/student/schedule)")
body(doc, "Giao diện học sinh gồm hai tab: Thời khoá biểu (chỉ đọc) và Đăng ký môn học.")

tbl(doc,
    ["Trạng thái hệ thống", "Học sinh có thể làm gì"],
    [
        ["Giai đoạn 2 chưa mở",  "Chỉ xem Tab Thời khoá biểu. Nút Đăng ký bị ẩn."],
        ["Giai đoạn 2 đang MỞ",  "Đăng ký và huỷ môn học. Badge sĩ số cập nhật tự động mỗi 5 giây."],
        ["Giai đoạn 2 đã ĐÓNG",  "Chỉ xem TKB cá nhân — không đăng ký, không huỷ, không tương tác."],
    ],
    widths=[5.0, 11.5]
)

H2(doc, "5.1. Tab Thời khoá biểu")
body(doc, "Hiển thị lịch tuần cá nhân của học sinh (chỉ đọc). Ô lịch tô màu = đã đăng ký môn học đó, kèm tên môn, GV, phòng học.")

H2(doc, "5.2. Tab Đăng ký môn học")
body(doc, "Gồm hai khu vực:")
bullet(doc, "Khu vực trên — Môn đã đăng ký: danh sách các lớp HS đã chọn, mỗi thẻ có nút Huỷ đăng ký.")
bullet(doc, "Khu vực dưới — Danh sách lớp có thể đăng ký: tất cả lớp phù hợp với khối của HS (đã có đủ địa điểm và sĩ số).")
body(doc, "Badge sĩ số trên mỗi thẻ lớp:")
bullet(doc, "Xanh lá = còn nhiều chỗ trống.")
bullet(doc, "Vàng = gần đầy (≥ 2/3 sĩ số đã đăng ký).")
bullet(doc, "Xám = đã đầy — không thể đăng ký thêm.")

H2(doc, "5.3. Đăng ký môn học")
note(doc, "Chỉ thực hiện được khi Giai đoạn 2 đang MỞ.", title="Điều kiện", bg="DEEAF1", tc=(0x1F, 0x49, 0x7D))
step(doc, 1, "Chuyển sang tab Đăng ký môn học.")
step(doc, 2, "Xem danh sách lớp ở khu vực dưới. Chú ý badge màu sĩ số.")
step(doc, 3, "Nhấn nút Đăng ký trên thẻ lớp muốn học.")
step(doc, 4, "Nếu không trùng lịch và còn chỗ: thông báo xanh xác nhận thành công.")
step(doc, 5, "Lớp vừa đăng ký xuất hiện ở khu vực trên và trong Tab Thời khoá biểu.")

H2(doc, "5.4. Huỷ đăng ký")
note(doc, "Chỉ thực hiện được khi Giai đoạn 2 đang MỞ.", title="Điều kiện", bg="DEEAF1", tc=(0x1F, 0x49, 0x7D))
body(doc, "Có hai cách:")
bullet(doc, "Từ khu vực trên: nhấn nút Huỷ đăng ký trên thẻ môn đã đăng ký → xác nhận.")
bullet(doc, "Từ danh sách lớp (khu vực dưới): tìm thẻ đang có nút Huỷ đăng ký màu đỏ → nhấn → xác nhận.")

H2(doc, "5.5. Xử lý trùng lịch")
body(doc, "Khi đăng ký một lớp có lịch trùng với lớp đã chọn trước:")
step(doc, 1, "Hệ thống dừng đăng ký và hiển thị modal thông báo trùng lịch.")
step(doc, 2, "Modal hiển thị tên môn, giáo viên và lịch của lớp đang bị trùng.")
step(doc, 3, "Nhấn Đóng để quay lại, sau đó chọn lớp khác phù hợp.")

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# 6. XỬ LÝ SỰ CỐ THƯỜNG GẶP
# ─────────────────────────────────────────────────────────────────────────────
H1(doc, "6. XỬ LÝ SỰ CỐ THƯỜNG GẶP")

tbl(doc,
    ["Sự cố", "Nguyên nhân thường gặp", "Cách xử lý"],
    [
        ["Sai mật khẩu khi đăng nhập",
         "Nhập sai hoặc quên mật khẩu",
         "Liên hệ Admin → khôi phục MK → nhận MK tạm 12 ký tự → đăng nhập → đổi MK mới"],
        ["GV không click được ô lịch",
         "Giai đoạn 1 chưa mở hoặc đã đóng",
         "Admin vào /admin/class-reg → nhấn Mở đăng ký GV"],
        ["HS không thấy nút Đăng ký",
         "GĐ2 chưa mở, hoặc lớp đã đầy (badge xám)",
         "Kiểm tra badge; nếu không xám → Admin mở Giai đoạn 2"],
        ["Lớp không hiện với HS",
         "Thiếu địa điểm hoặc sĩ số tối đa",
         "Admin vào /admin/class-reg → điền đủ phòng học & sĩ số cho lớp đó"],
        ["Lỗi khi nhập file Excel",
         "Sai định dạng cột hoặc CCCD/Mã GV trùng",
         "Tải lại file mẫu mới nhất; không đổi tên cột; kiểm tra dữ liệu trùng lặp"],
        ["Trang tải chậm hoặc báo lỗi",
         "Sự cố mạng hoặc máy chủ",
         "Nhấn F5 tải lại trang; báo bộ phận kỹ thuật nếu lỗi kéo dài hơn 5 phút"],
    ],
    widths=[4.5, 5.0, 7.0]
)

# ─────────────────────────────────────────────────────────────────────────────
# 7. QUY TRÌNH TỔNG QUÁT
# ─────────────────────────────────────────────────────────────────────────────
H1(doc, "7. QUY TRÌNH VẬN HÀNH TỔNG QUÁT")
body(doc, "Dưới đây là trình tự đúng để vận hành một chu kỳ đăng ký đầy đủ:")

steps_flow = [
    "Admin tạo tài khoản GV và HS (thủ công hoặc nhập Excel).",
    "Admin vào /admin/class-reg → nhấn Mở đăng ký GV (bắt đầu Giai đoạn 1).",
    "Giáo viên đăng nhập → đăng ký lịch mở lớp trên bảng thời khoá biểu.",
    "Admin nhấn Đóng đăng ký GV (kết thúc Giai đoạn 1).",
    "Admin điền phòng học và sĩ số tối đa cho từng lớp tại /admin/class-reg.",
    "Kiểm tra: không còn lớp nào có viền đỏ cảnh báo.",
    "Admin vào /admin/enrollment → nhấn Mở đăng ký HS (bắt đầu Giai đoạn 2).",
    "Học sinh đăng nhập → đăng ký các môn học theo khối.",
    "Admin nhấn Đóng đăng ký HS (kết thúc Giai đoạn 2).",
    "Admin xuất file Excel kết quả tại /admin/enrollment → dùng để lập danh sách lớp, báo cáo.",
]
for i, s in enumerate(steps_flow, 1):
    step(doc, i, s)

doc.add_paragraph()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("─" * 45), size=10, color=(0x80, 0x80, 0x80))

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(p.add_run("Tài liệu thuộc Trường THPT Nguyễn Hữu Cầu  ·  mynhc v1.1  ·  Năm học 2026–2027"),
         italic=True, size=10, color=(0x60, 0x60, 0x60))

# ─────────────────────────────────────────────────────────────────────────────
add_header_footer(doc)
# ─────────────────────────────────────────────────────────────────────────────

output_path = "/Users/ddang/Documents/Subject_Survey/ClassReg/Hướng_dẫn_sử_dụng_mynhc.docx"
doc.save(output_path)
print(f"Đã lưu: {output_path}")
