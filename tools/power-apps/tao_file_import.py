"""Tạo file Excel mẫu để import vào SharePoint Lists (mỗi sheet = 1 Excel Table).

Chạy: python tao_file_import.py  -> mau_import_sharepoint.xlsx
Dữ liệu là GIẢ LẬP; tham số chính sách là GIẢ ĐỊNH mẫu.
"""
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo

VUNG = ["Miền Bắc", "Miền Trung", "Miền Nam"]
CAP = ["Đại lý", "Trưởng nhóm (UM)", "Trưởng ban (BM)"]
HO = ["Nguyễn", "Trần", "Lê", "Phạm", "Hoàng", "Võ", "Đặng", "Bùi"]
TEN = ["An", "Bình", "Chi", "Dũng", "Hà", "Khánh", "Lan", "Minh", "Nam", "Thảo"]


def add_table(wb, name, headers, rows, fmts, widths, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = name
    ws.append(headers)
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows():
        for c in row:
            c.font = Font(name="Arial", bold=c.row == 1)
            f = fmts.get(c.column)
            if f and c.row > 1:
                c.number_format = f
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    t = Table(displayName=f"tbl{name}", ref=ws.dimensions)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)


def main():
    random.seed(7)
    agents = [(f"DL{i:04d}", f"{random.choice(HO)} {random.choice(TEN)}",
               random.choices(CAP, weights=[7, 2, 1])[0], random.choice(VUNG))
              for i in range(1, 31)]
    rows = []
    for ky in ["2026-09", "2026-10"]:
        for ma, ten, cap, vung in agents:
            fyp = random.choice([0, 8, 15, 25, 40, 60, 90, 150]) * 1_000_000
            rows.append([ma, ten, cap, vung, fyp, 0 if fyp == 0 else random.randint(1, 6),
                         random.choice([0.6, 0.7, 0.75, 0.8, 0.85, 0.9, 0.95]),
                         random.choice([0, 0, 0, 500_000, 1_000_000]), ky])

    wb = Workbook()
    add_table(wb, "DaiLy",
              ["MaDL", "HoTen", "CapBac", "Vung", "FYP", "SoHD", "K2", "ThuNhapKhac", "Ky"],
              rows, {5: "#,##0", 7: "0%", 8: "#,##0"},
              [10, 20, 18, 13, 15, 8, 8, 14, 10], first=True)
    add_table(wb, "ChinhSach", ["Ten", "TyLeFYC", "K2Min", "HDMin", "Thue"],
              [["Mặc định", 0.25, 0.75, 2, 0.05]],
              {2: "0.0%", 3: "0%", 5: "0.0%"}, [14, 10, 10, 10, 10])
    add_table(wb, "BacThuong", ["Bac", "FYCTu", "TyLe"],
              [["Bậc 0", 0, 0], ["Bậc 1", 5_000_000, 0.10],
               ["Bậc 2", 10_000_000, 0.20], ["Bậc 3", 20_000_000, 0.30]],
              {2: "#,##0", 3: "0%"}, [10, 14, 10])
    add_table(wb, "PhanQuyen", ["Email", "VaiTro", "Vung"],
              [["admin@congty.vn", "Admin", ""], ["ad.mienbac@congty.vn", "AD", "Miền Bắc"]],
              {}, [28, 10, 14])
    out = Path(__file__).with_name("mau_import_sharepoint.xlsx")
    wb.save(out)
    print("Đã tạo", out)


if __name__ == "__main__":
    main()
