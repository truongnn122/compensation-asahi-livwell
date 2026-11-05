"""Tạo báo cáo thu nhập Đại lý từ file Excel đầu vào.

Chạy:
    python bao_cao_thu_nhap.py dau_vao_mau.xlsx --ky "T10-2026"
Kết quả: bao_cao_thu_nhap_T10-2026.xlsx gồm 3 sheet
    TongHop   – KPI, tổng hợp theo Vùng / Cấp bậc, Top 10, biểu đồ
    ChiTiet   – thu nhập từng Đại lý (toàn bộ là công thức Excel, sửa được)
    ChinhSach – tham số chính sách dùng cho kỳ này
"""
import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = "Arial"
F = Font(name=FONT)
B = Font(name=FONT, bold=True)
GREEN = Font(name=FONT, color="008000")
HEAD_FONT = Font(name=FONT, bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
VND = "#,##0;(#,##0);-"
PCT = "0.0%;(0.0%);-"

COT_BAT_BUOC = ["Mã ĐL", "Họ tên", "Cấp bậc", "Vùng", "FYP (VNĐ)",
                "Số HĐ", "K2 (%)", "Thu nhập khác (VNĐ)"]
CS = "ChinhSach"  # tên sheet chính sách


# ---------- đọc & kiểm tra đầu vào ----------
def doc_dau_vao(path):
    df = pd.read_excel(path, sheet_name="DuLieuDaiLy", usecols=range(8))
    thieu = [c for c in COT_BAT_BUOC if c not in df.columns]
    if thieu:
        sys.exit(f"Thiếu cột trong DuLieuDaiLy: {thieu}")
    df = df.dropna(subset=["Mã ĐL"])
    for c in ["FYP (VNĐ)", "Số HĐ", "K2 (%)", "Thu nhập khác (VNĐ)"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    trung = df[df["Mã ĐL"].duplicated()]["Mã ĐL"].tolist()
    if trung:
        sys.exit(f"Mã ĐL bị trùng: {trung}")
    if (df["K2 (%)"] > 1).any():
        print("Cảnh báo: có K2 > 100% – kiểm tra lại đã nhập dạng % chưa.")

    cs = pd.read_excel(path, sheet_name=CS, header=None)
    tham_so = {
        "fyc_rate": float(cs.iat[2, 1]),
        "k2_min": float(cs.iat[3, 1]),
        "hd_min": float(cs.iat[4, 1]),
        "thue": float(cs.iat[5, 1]),
        "bac": [(str(cs.iat[r, 0]), float(cs.iat[r, 1]), float(cs.iat[r, 2]))
                for r in range(9, 13)],
    }
    return df, tham_so


# ---------- helpers ----------
def header(ws, row, names, widths=None):
    for i, n in enumerate(names, 1):
        c = ws.cell(row, i, n)
        c.font, c.fill, c.border = HEAD_FONT, HEAD_FILL, THIN
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if widths:
            ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.row_dimensions[row].height = 32


def put(ws, r, c, v, fmt=None, font=F, fill=None):
    cell = ws.cell(r, c, v)
    cell.font, cell.border = font, THIN
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell


# ---------- sheet ChinhSach ----------
def ghi_chinh_sach(wb, ts):
    ws = wb.create_sheet(CS)
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws["A1"] = "THAM SỐ CHÍNH SÁCH ÁP DỤNG KỲ NÀY"
    ws["A1"].font = Font(name=FONT, bold=True, size=12, color="1F4E78")
    blue = Font(name=FONT, color="0000FF")
    for r, (lab, key, fmt) in enumerate([
        ("Tỷ lệ hoa hồng năm đầu (FYC/FYP)", "fyc_rate", "0.0%"),
        ("K2 tối thiểu để nhận thưởng", "k2_min", "0%"),
        ("Số HĐ tối thiểu để nhận thưởng", "hd_min", "0"),
        ("Thuế TNCN khấu trừ", "thue", "0.0%"),
    ], 3):
        put(ws, r, 1, lab)
        put(ws, r, 2, ts[key], fmt, blue)
    header_row = 9
    for i, n in enumerate(["Bậc", "FYC từ (VNĐ)", "Tỷ lệ thưởng / FYC"], 1):
        c = ws.cell(header_row, i, n)
        c.font, c.fill, c.border = HEAD_FONT, HEAD_FILL, THIN
    for r, (bac, tu, tl) in enumerate(ts["bac"], 10):
        put(ws, r, 1, bac)
        put(ws, r, 2, tu, VND, blue)
        put(ws, r, 3, tl, "0%", blue)
    ws["A15"] = "Nguồn: sheet ChinhSach của file đầu vào. Chữ xanh = tham số, sửa tại đây sẽ tính lại toàn bộ báo cáo."
    ws["A15"].font = Font(name=FONT, italic=True, color="7F7F7F")


# ---------- sheet ChiTiet ----------
CT_COLS = COT_BAT_BUOC + ["FYC (VNĐ)", "Đạt ĐK thưởng", "Tỷ lệ thưởng",
                          "Thưởng HĐ (VNĐ)", "Tổng TN trước thuế", "Thuế TNCN",
                          "Thực nhận (VNĐ)"]
CT_W = [10, 22, 18, 13, 15, 8, 8, 15, 15, 11, 10, 15, 16, 13, 16]
FIRST = 5  # dòng dữ liệu đầu tiên


def ghi_chi_tiet(wb, df, ky):
    ws = wb.create_sheet("ChiTiet")
    ws["A1"] = f"CHI TIẾT THU NHẬP ĐẠI LÝ – KỲ {ky}"
    ws["A1"].font = Font(name=FONT, bold=True, size=13, color="1F4E78")
    ws["A2"] = "Cột I–O là công thức tham chiếu sheet ChinhSach."
    ws["A2"].font = Font(name=FONT, italic=True, color="7F7F7F")
    header(ws, 4, CT_COLS, CT_W)

    for i, rec in enumerate(df.itertuples(index=False)):
        r = FIRST + i
        vals = list(rec)
        fmts = [None, None, None, None, VND, "0", "0%", VND]
        for c, (v, f) in enumerate(zip(vals, fmts), 1):
            put(ws, r, c, v, f)
        put(ws, r, 9, f"=E{r}*{CS}!$B$3", VND)
        put(ws, r, 10, f'=IF(AND(F{r}>={CS}!$B$5,G{r}>={CS}!$B$4),"Đạt","Không")')
        put(ws, r, 11, f'=IF(J{r}="Đạt",INDEX({CS}!$C$10:$C$13,'
                       f'MATCH(I{r},{CS}!$B$10:$B$13,1)),0)', "0%")
        put(ws, r, 12, f"=I{r}*K{r}", VND)
        put(ws, r, 13, f"=I{r}+L{r}+H{r}", VND)
        put(ws, r, 14, f"=M{r}*{CS}!$B$6", VND)
        put(ws, r, 15, f"=M{r}-N{r}", VND)

    last = FIRST + len(df) - 1
    t = last + 1
    put(ws, t, 1, "TỔNG", font=B, fill=TOTAL_FILL)
    for c in range(2, 16):
        col = get_column_letter(c)
        if c in (5, 6, 8, 9, 12, 13, 14, 15):
            put(ws, t, c, f"=SUM({col}{FIRST}:{col}{last})",
                "0" if c == 6 else VND, B, TOTAL_FILL)
        elif c == 10:
            put(ws, t, c, f'=COUNTIF(J{FIRST}:J{last},"Đạt")', "0", B, TOTAL_FILL)
        else:
            put(ws, t, c, None, fill=TOTAL_FILL)
    ws.freeze_panes = f"C{FIRST}"
    ws.auto_filter.ref = f"A4:O{last}"
    return last


# ---------- sheet TongHop ----------
def ghi_tong_hop(wb, df, ts, ky, last):
    ws = wb.active
    ws.title = "TongHop"
    for col, w in zip("ABCDEFGH", [26, 14, 14, 18, 18, 18, 14, 4]):
        ws.column_dimensions[col].width = w
    ws["A1"] = f"BÁO CÁO THU NHẬP ĐẠI LÝ – KỲ {ky}"
    ws["A1"].font = Font(name=FONT, bold=True, size=14, color="1F4E78")

    rng = lambda c: f"ChiTiet!${c}${FIRST}:${c}${last}"

    # KPI
    ws["A3"] = "CHỈ SỐ CHÍNH"
    ws["A3"].font = B
    kpi = [
        ("Tổng số Đại lý", f"=COUNTA({rng('A')})", "0"),
        ("Đại lý hoạt động (FYP>0)", f'=COUNTIF({rng("E")},">0")', "0"),
        ("Tỷ lệ hoạt động", "=IFERROR(B5/B4,0)", PCT),
        ("Đại lý đạt ĐK thưởng", f'=COUNTIF({rng("J")},"Đạt")', "0"),
        ("Tổng FYP (VNĐ)", f"=SUM({rng('E')})", VND),
        ("Tổng thu nhập trước thuế (VNĐ)", f"=SUM({rng('M')})", VND),
        ("Tổng thực nhận (VNĐ)", f"=SUM({rng('O')})", VND),
        ("Chi phí thu nhập / FYP", "=IFERROR(B9/B8,0)", PCT),
        ("TN bình quân / ĐL hoạt động", "=IFERROR(B9/B5,0)", VND),
    ]
    for r, (lab, fml, fmt) in enumerate(kpi, 4):
        put(ws, r, 1, lab)
        put(ws, r, 2, fml, fmt, GREEN)

    # Tổng hợp theo nhóm
    def bang(start, title, key_col, keys):
        ws.cell(start, 1, title).font = B
        header(ws, start + 1, ["", "Số ĐL", "ĐL hoạt động", "FYP (VNĐ)",
                               "FYC (VNĐ)", "Thưởng HĐ (VNĐ)", "Thực nhận (VNĐ)"])
        ws.cell(start + 1, 1, title.split("THEO ")[-1].title())
        r = start + 2
        for k in keys:
            crit = f"{rng(key_col)},$A{r}"
            put(ws, r, 1, k)
            put(ws, r, 2, f"=COUNTIF({crit})", "0", GREEN)
            put(ws, r, 3, f'=COUNTIFS({crit},{rng("E")},">0")', "0", GREEN)
            for c, src in zip(range(4, 8), "EILO"):
                put(ws, r, c, f"=SUMIFS({rng(src)},{crit})", VND, GREEN)
            r += 1
        put(ws, r, 1, "Tổng", font=B, fill=TOTAL_FILL)
        for c in range(2, 8):
            col = get_column_letter(c)
            put(ws, r, c, f"=SUM({col}{start + 2}:{col}{r - 1})",
                "0" if c < 4 else VND, B, TOTAL_FILL)
        return start + 2, r - 1, r

    vung = sorted(df["Vùng"].dropna().unique())
    cap = sorted(df["Cấp bậc"].dropna().unique())
    v_first, v_last, end1 = bang(15, "TỔNG HỢP THEO VÙNG", "D", vung)
    _, _, end2 = bang(end1 + 3, "TỔNG HỢP THEO CẤP BẬC", "C", cap)

    # Top 10 – thứ tự xếp bằng Python, giá trị lấy bằng công thức
    top_start = end2 + 3
    ws.cell(top_start, 1, "TOP 10 ĐẠI LÝ THEO THỰC NHẬN").font = B
    header(ws, top_start + 1, ["Họ tên", "Mã ĐL", "Vùng", "FYP (VNĐ)",
                               "FYC (VNĐ)", "Thưởng HĐ (VNĐ)", "Thực nhận (VNĐ)"])
    fyc = df["FYP (VNĐ)"] * ts["fyc_rate"]
    du_dk = (df["Số HĐ"] >= ts["hd_min"]) & (df["K2 (%)"] >= ts["k2_min"])
    tl = fyc.apply(lambda x: max(t for _, tu, t in ts["bac"] if x >= tu)
                   if x >= 0 else 0) * du_dk
    thuc_nhan = (fyc * (1 + tl) + df["Thu nhập khác (VNĐ)"]) * (1 - ts["thue"])
    order = thuc_nhan.sort_values(ascending=False).index[:10]
    for i, idx in enumerate(order):
        r, src = top_start + 2 + i, FIRST + df.index.get_loc(idx)
        for c, (col, fmt) in enumerate(zip("BADEILO", [None, None, None, VND, VND, VND, VND]), 1):
            put(ws, r, c, f"=ChiTiet!{col}{src}", fmt, GREEN)
    ws.cell(top_start + 13, 1,
            "Ghi chú: thứ tự Top 10 được xếp lúc chạy tool; nếu sửa ChinhSach trong file này, "
            "hãy chạy lại tool để xếp hạng lại.").font = Font(name=FONT, italic=True, color="7F7F7F")

    # Biểu đồ thực nhận theo vùng
    ch = BarChart()
    ch.title = "Thực nhận theo Vùng (VNĐ)"
    ch.y_axis.title = "VNĐ"
    ch.y_axis.numFmt = "#,##0"
    ch.legend = None
    ch.add_data(Reference(ws, min_col=7, min_row=v_first - 1, max_row=v_last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=v_first, max_row=v_last))
    ch.height, ch.width = 7.5, 13
    ws.add_chart(ch, "D3")


def main():
    ap = argparse.ArgumentParser(description="Tạo báo cáo thu nhập Đại lý")
    ap.add_argument("input", help="File Excel đầu vào (sheet DuLieuDaiLy, ChinhSach)")
    ap.add_argument("--ky", default="KY", help='Kỳ báo cáo, vd "T10-2026"')
    ap.add_argument("-o", "--output", help="Tên file kết quả")
    a = ap.parse_args()

    df, ts = doc_dau_vao(a.input)
    wb = Workbook()
    last = ghi_chi_tiet(wb, df, a.ky)
    ghi_tong_hop(wb, df, ts, a.ky, last)
    ghi_chinh_sach(wb, ts)
    out = Path(a.output or f"bao_cao_thu_nhap_{a.ky}.xlsx")
    wb.save(out)
    print(f"Đã tạo {out} – {len(df)} Đại lý")


if __name__ == "__main__":
    main()
