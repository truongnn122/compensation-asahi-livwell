# Hướng dẫn dựng App báo cáo thu nhập Đại lý bằng Power Apps

Mục tiêu: người xem mở app (web / Teams / điện thoại) là thấy số kỳ mới nhất. Mỗi kỳ chỉ cần dán dữ liệu mới vào SharePoint List, không chạy tool, không gửi lại file.

```
SharePoint Lists (DaiLy, ChinhSach, BacThuong, PhanQuyen)
          │  đọc
          ▼
Power Apps: tính FYC, thưởng, thuế → KPI, bảng, biểu đồ, Top 10
```

> Tham số chính sách trong file mẫu (FYC 25%, bậc thưởng, K2 ≥ 75%, thuế 5%) là **giả định**, cần thay bằng chính sách thật.
> Phù hợp khi số Đại lý mỗi kỳ < 2.000 dòng.

---

## Bước 0 – Kiểm tra quyền
1. Vào **make.powerapps.com** bằng tài khoản công ty.
2. Vào được và bấm **+ Create** được là có quyền dùng. Nếu báo lỗi license thì liên hệ IT.

## Bước 1 – Tạo SharePoint Lists từ file mẫu
File `mau_import_sharepoint.xlsx` có 4 sheet, mỗi sheet là 1 Excel Table.

Với từng sheet: vào trang SharePoint của phòng → **+ New → List → From Excel** → chọn file → chọn table → kiểm tra kiểu cột → **Create**.

| List | Cột | Kiểu |
|---|---|---|
| **DaiLy** | MaDL, HoTen, CapBac, Vung, Ky | Single line of text |
| | FYP, SoHD, K2, ThuNhapKhac | Number |
| **ChinhSach** (1 dòng) | Ten | Text |
| | TyLeFYC, K2Min, HDMin, Thue | Number |
| **BacThuong** | Bac | Text |
| | FYCTu, TyLe | Number |
| **PhanQuyen** | Email, VaiTro (Admin/AD), Vung | Text |

Lưu ý:
- **Giữ tên cột không dấu** như file mẫu vì Power Apps dùng tên gốc (internal name). Muốn hiển thị có dấu thì đổi tên *sau khi* tạo list.
- `K2`, tỷ lệ nhập dạng số thập phân: 85% = `0.85`.
- `Ky` dạng `YYYY-MM` (vd `2026-10`) để sắp xếp đúng.

## Bước 2 – Tạo app
1. make.powerapps.com → **+ Create → Blank app → Blank canvas app** → Format **Tablet**.
2. **Data → Add data → SharePoint** → chọn site → tick 4 list trên.

### App.OnStart
```powerfx
Set(varUser, LookUp(PhanQuyen, Lower(Email) = Lower(User().Email)));
Set(varCS, First(ChinhSach));
ClearCollect(colBac, SortByColumns(BacThuong, "FYCTu", SortOrder.Descending));
ClearCollect(colKy, Sort(Distinct(DaiLy, Ky), Value, SortOrder.Descending))
```

### Màn hình Dashboard
**Dropdown `drpKy`** (chọn kỳ)
- Items: `colKy`
- Default: `First(colKy).Value`
- OnChange: `Select(btnTinh)`

**Button `btnTinh`** (ẩn đi: Visible = `false`) – OnSelect:
```powerfx
ClearCollect(colTN,
  ForAll(
    Filter(DaiLy, Ky = drpKy.Selected.Value) As d,
    With(
      { fyc: d.FYP * varCS.TyLeFYC,
        dat: d.SoHD >= varCS.HDMin && d.K2 >= varCS.K2Min },
      With(
        { tl: If(dat, Coalesce(LookUp(colBac, FYCTu <= fyc).TyLe, 0), 0) },
        With(
          { tong: fyc + fyc * tl + d.ThuNhapKhac },
          { MaDL: d.MaDL, HoTen: d.HoTen, CapBac: d.CapBac, Vung: d.Vung,
            FYP: d.FYP, SoHD: d.SoHD, K2: d.K2, FYC: fyc,
            DatDK: If(dat, "Đạt", "Không"), TyLeThuong: tl, Thuong: fyc * tl,
            TongTruocThue: tong, Thue: tong * varCS.Thue,
            ThucNhan: tong * (1 - varCS.Thue) }
        )
      )
    )
  )
);
// Phân quyền: AD chỉ thấy vùng của mình
If(varUser.VaiTro <> "Admin", RemoveIf(colTN, Vung <> varUser.Vung))
```
Screen **OnVisible**: `Select(btnTinh)`

> `colBac` đã sắp xếp giảm dần, nên `LookUp(colBac, FYCTu <= fyc)` trả về bậc cao nhất mà Đại lý đạt.

**KPI** (Label, Text):
| Chỉ số | Công thức |
|---|---|
| Tổng số ĐL | `CountRows(colTN)` |
| ĐL hoạt động | `CountIf(colTN, FYP > 0)` |
| Tỷ lệ hoạt động | `Text(IfError(CountIf(colTN, FYP > 0) / CountRows(colTN), 0), "0.0%")` |
| ĐL đạt ĐK thưởng | `CountIf(colTN, DatDK = "Đạt")` |
| Tổng FYP | `Text(Sum(colTN, FYP), "#,##0")` |
| Tổng thực nhận | `Text(Sum(colTN, ThucNhan), "#,##0")` |
| Chi phí TN / FYP | `Text(IfError(Sum(colTN, TongTruocThue) / Sum(colTN, FYP), 0), "0.0%")` |

**Bảng tổng hợp theo Vùng** – Data table, Items:
```powerfx
AddColumns(
  GroupBy(colTN, Vung, Nhom),
  SoDL, CountRows(Nhom),
  HoatDong, CountIf(Nhom, FYP > 0),
  TongFYP, Sum(Nhom, FYP),
  TongThuong, Sum(Nhom, Thuong),
  TongThucNhan, Sum(Nhom, ThucNhan)
)
```
Theo Cấp bậc: thay `Vung` bằng `CapBac`.

**Biểu đồ** – Insert → Charts → **Column chart**: Items = bảng theo Vùng ở trên (bỏ cột `Nhom`):
```powerfx
ShowColumns(<công thức bảng theo Vùng>, Vung, TongThucNhan)
```

**Top 10** – Data table, Items:
```powerfx
FirstN(Sort(colTN, ThucNhan, SortOrder.Descending), 10)
```

### Màn hình Chi tiết
- **Text input `txtTim`** + **Data table**, Items:
```powerfx
Sort(
  Filter(colTN, IsBlank(txtTim.Text) || txtTim.Text in HoTen || txtTim.Text in MaDL),
  ThucNhan, SortOrder.Descending
)
```
- Định dạng số trong Data table: chọn cột → thêm format bằng `Text(ThisItem.ThucNhan, "#,##0")` nếu dùng Gallery.

## Bước 3 – Chia sẻ
1. **Save → Publish**.
2. **Share** → nhập email/nhóm người xem.
3. Người xem cần quyền **Read** trên 4 SharePoint List (Share app sẽ gợi ý).
4. Gắn vào Teams: trong Teams → **Apps → Built for your org** hoặc thêm tab Power Apps vào kênh.

## Cập nhật hằng kỳ
1. Mở list **DaiLy** → **Edit in grid view**.
2. Dán dữ liệu kỳ mới từ Excel (đúng thứ tự cột, cột `Ky` = kỳ mới) → **Exit grid view**.
3. Xong. Người xem mở app là thấy kỳ mới (mặc định chọn kỳ gần nhất).

Đổi chính sách: sửa dòng trong list **ChinhSach** / **BacThuong**.

## Lưu ý bảo mật
- Phân quyền theo vùng trong app **chỉ ẩn trên giao diện**. Ai có quyền đọc list `DaiLy` vẫn mở trực tiếp list trên SharePoint được.
- Nếu AD tuyệt đối không được xem vùng khác: tách list theo vùng (DaiLy_MienBac, …) và phân quyền SharePoint từng list, hoặc dùng quyền theo từng item.
- Ẩn list khỏi menu site: List settings → Advanced → không hiển thị trên Navigation.

## Lỗi thường gặp
| Hiện tượng | Cách xử lý |
|---|---|
| Cảnh báo tam giác vàng "delegation" | Bình thường khi dữ liệu ít (< 2.000 dòng). Vào Settings → General → Data row limit đặt **2000** |
| `GroupBy` báo lỗi | Bản Power Apps cũ: dùng `GroupBy(colTN, "Vung", "Nhom")` và `AddColumns(..., "SoDL", CountRows(Nhom), ...)` |
| Số thưởng = 0 hết | Kiểm tra `K2` nhập dạng 0.85 chứ không phải 85; kiểm tra `colBac` có dữ liệu |
| Không thấy kỳ mới | Bấm Refresh hoặc mở lại app (OnStart chạy lại) |
