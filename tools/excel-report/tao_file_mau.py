"""Tạo file Excel đầu vào mẫu: sheet DuLieuDaiLy + sheet ChinhSach.

Chạy:  python tao_file_mau.py            -> tạo dau_vao_mau.xlsx
Ô nền vàng / chữ xanh = ô cần nhập. Số chính sách là GIẢ ĐỊNH để chạy thử.
"""
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FONT = "Arial"
INPUT_FONT = Font(name=FONT, color="0000FF")
HEAD_FONT = Font(name=FONT, bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Border(*(Side(style="thin", color="BFBFBF"),) * 4)

VUNG = ["Miền Bắc", "Miền Trung", "Miền Nam"]
CAP = ["Đại lý", "Trưởng nhóm (UM)", "Trưởng ban (BM)"]
HO = ["Nguyễn", "Trần", "Lê", "Phạm", "Hoàng", "Võ", "Đặng", "Bùi"]
TEN = ["An", "Bình", "Chi", "Dũng", "Hà", "Khánh", "Lan", "Minh", "Nam", "Thảo"]

DATA_COLS = [
    ("Mã ĐL", 12), ("Họ tên", 22), ("Cấp bậc", 18), ("Vùng", 14),
    ("FYP (VNĐ)", 16), ("Số HĐ", 9), ("K2 (%)", 10), ("Thu nhập khác (VNĐ)", 18),
]


def head(ws, row, cols):
    for i, (name, width) in enumerate(cols, 1):
        c = ws.cell(row, i, name)
        c.font, c.fill, c.border = HEAD_FONT, HEAD_FILL, THIN
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[c.column_letter].width = width


def sheet_du_lieu(wb):
    ws = wb.active
    ws.title = "DuLieuDaiLy"
    head(ws, 1, DATA_COLS)
    random.seed(7)
    for r in range(2, 32):
        fyp = random.choice([0, 8, 15, 25, 40, 60, 90, 150]) * 1_000_000
        row = [
            f"DL{r - 1:04d}",
            f"{random.choice(HO)} {random.choice(TEN)}",
            random.choices(CAP, weights=[7, 2, 1])[0],
            random.choice(VUNG),
            fyp,
            0 if fyp == 0 else random.randint(1, 6),
            random.choice([0.6, 0.7, 0.75, 0.8, 0.85, 0.9, 0.95]),
            random.choice([0, 0, 0, 500_000, 1_000_000]),
        ]
        for i, v in enumerate(row, 1):
            c = ws.cell(r, i, v)
            c.font, c.border = INPUT_FONT, THIN
            if i in (5, 8):
                c.number_format = "#,##0"
            if i == 7:
                c.number_format = "0%"
    ws.freeze_panes = "A2"
    ws["J1"] = "Hướng dẫn"
    ws["J1"].font = Font(name=FONT, bold=True)
    for i, t in enumerate([
        "- Mỗi dòng = 1 Đại lý trong kỳ. Dữ liệu dòng 2–31 là MẪU, xoá và dán dữ liệu thật.",
        "- K2 nhập dạng % (vd 85%). FYP & thu nhập khác nhập theo VNĐ.",
        "- Không đổi tên cột / tên sheet.",
    ], 2):
        ws.cell(i, 10, t).font = Font(name=FONT)


def sheet_chinh_sach(wb):
    ws = wb.create_sheet("ChinhSach")
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 55

    def title(r, text):
        ws.cell(r, 1, text).font = Font(name=FONT, bold=True, size=12, color="1F4E78")

    def param(r, label, value, fmt, note):
        ws.cell(r, 1, label).font = Font(name=FONT)
        c = ws.cell(r, 2, value)
        c.font, c.fill, c.number_format, c.border = INPUT_FONT, INPUT_FILL, fmt, THIN
        ws.cell(r, 4, note).font = Font(name=FONT, italic=True, color="7F7F7F")

    title(1, "THAM SỐ CHÍNH SÁCH (giả định mẫu – thay bằng chính sách thật)")
    param(3, "Tỷ lệ hoa hồng năm đầu (FYC/FYP)", 0.25, "0.0%", "Giả định 1 tỷ lệ chung; thực tế khác theo sản phẩm/kỳ hạn")
    param(4, "K2 tối thiểu để nhận thưởng", 0.75, "0%", "Giả định")
    param(5, "Số HĐ tối thiểu để nhận thưởng", 2, "0", "Giả định")
    param(6, "Thuế TNCN khấu trừ", 0.05, "0.0%", "Cần xác nhận với Kế toán (TT 40/2021: ngưỡng doanh thu năm)")

    title(8, "BẬC THƯỞNG HOẠT ĐỘNG (theo FYC, sắp xếp tăng dần)")
    for i, h in enumerate(["FYC từ (VNĐ)", "Tỷ lệ thưởng / FYC"], 2):
        c = ws.cell(9, i, h)
        c.font, c.fill, c.border = HEAD_FONT, HEAD_FILL, THIN
    ws.cell(9, 1, "Bậc").font = Font(name=FONT, bold=True)
    for r, (bac, tu, tl) in enumerate([
        ("Bậc 0", 0, 0), ("Bậc 1", 5_000_000, 0.10),
        ("Bậc 2", 10_000_000, 0.20), ("Bậc 3", 20_000_000, 0.30),
    ], 10):
        ws.cell(r, 1, bac).font = Font(name=FONT)
        for col, v, fmt in ((2, tu, "#,##0"), (3, tl, "0%")):
            c = ws.cell(r, col, v)
            c.font, c.fill, c.number_format, c.border = INPUT_FONT, INPUT_FILL, fmt, THIN
    ws["D10"] = "Giữ đúng 4 dòng B10:C13; bậc đầu tiên phải bắt đầu từ 0"
    ws["D10"].font = Font(name=FONT, italic=True, color="7F7F7F")
    ws["B3"].comment = Comment("Giả định mẫu do Claude đặt, chưa phải chính sách thật", "tool")


def main(out="dau_vao_mau.xlsx"):
    wb = Workbook()
    sheet_du_lieu(wb)
    sheet_chinh_sach(wb)
    path = Path(__file__).with_name(out)
    wb.save(path)
    print("Đã tạo", path)


if __name__ == "__main__":
    main()
