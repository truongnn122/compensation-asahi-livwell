# Tool tính Phụ cấp cố định (FA) cho PM/PD – Microsoft Access

Căn cứ: Thông báo **ALL/AD/2210-006** (22/10/2026), Phụ lục 1 mục **A.I**. Hệ số ngày làm việc theo ALL/AD/2210-005 mục D.z.

## 1. Vị trí trong mô hình

| File | Vai trò |
|---|---|
| `Data.accdb` | Dữ liệu dùng chung: `Tbl_Agent`, `Tbl_Attendance`. **Dùng chung với tool FA LP**, không tạo thêm |
| **`FA_PM_App.accdb`** (file này) | Tool + dữ liệu riêng của FA PM/PD: tham số, `tblFAPM_Input`, **`tblFAPM_Result`**, trạng thái chốt kỳ |
| `TongHop.accdb` | Tự nhận diện file này khi bấm **Thêm file khoản thưởng** (mã `FA_PM`) |

Tool này chỉ tính cho agent có chức danh **PM** hoặc **PD**. Tool FA LP chỉ tính cho **LP**.

## 2. Cài đặt

1. Chép `Build_FAPM_Tool.vbs`, `modFAPM.bas`, `modUpdate.bas` vào **cùng thư mục với `Data.accdb`** đang dùng (thư mục của tool FA LP).
2. Double-click `Build_FAPM_Tool.vbs`. Script tạo `FA_PM_App.accdb` và kết nối vào `Data.accdb` có sẵn. Nếu thư mục chưa có `Data.accdb`, script tạo mới.
3. Mở `FA_PM_App.accdb` → **Enable Content**. Danh sách agent và chấm công đã nạp ở tool LP dùng được ngay, không phải nạp lại.

Nâng cấp sau này: chép `modFAPM.bas` mới vào thư mục → bấm **Cập nhật tool**.

## 3. Dữ liệu cần nhập hàng tháng (sheet `FAPM_Input`)

Mỗi dòng là 1 PM/PD cho 1 tháng/kỳ, số liệu **lũy kế đến ngày chốt ACK của kỳ**:

| Cột | Nội dung |
|---|---|
| `Period`, `Ky`, `AgentCode` | Tháng (YYYYMM), kỳ 1–4, mã PM/PD |
| `GroupFYC` | FYC nhóm trực tiếp phát sinh trong tháng (VND). Nhóm trực tiếp = cá nhân PM + các LP báo cáo trực tiếp (Phụ lục 2 văn bản 005) |
| `ActiveLP` | Số LP nhóm trực tiếp **hoạt động** trong tháng: LP có ≥ 1 HĐ và ≥ 4 triệu FYC phát sinh |
| `K2Pct` | %K2 nhóm trực tiếp (ví dụ 85). Để trống nếu chưa có K2 |
| `StatusEOM` | Tình trạng cuối tháng (Active/SE). Trống thì lấy `AGENT_STATUS` trong Tbl_Agent |

Số ngày đi làm lấy từ `Tbl_Attendance`. Ngày thăng tiến lấy từ cột `PROMOTION_DATE` trong Tbl_Agent; trống thì dùng `APPOINTMENT_DATE`.

Khi anh/chị gửi bảng doanh số theo hợp đồng và sơ đồ nhóm, tool sẽ tự tính `GroupFYC` và `ActiveLP`, không phải nhập tay.

## 4. Công thức đã cài

`FA = Phụ cấp/ngày × Hệ số ngày làm việc × %Hệ số K2 × %Hệ số Lượt LP hoạt động`

- **Phụ cấp/ngày:** 0,3 / 0,5 / 0,8 triệu khi FYC nhóm ≥ 10 / 30 / 60 triệu **và** ≥ 2 LP hoạt động. Không đạt thì 0.
- **%Hệ số K2:** ≥80% → 120%; ≥70% → 110%; ≥60% hoặc N/A → 100%; <60% → 50%.
- **%Hệ số Lượt LP hoạt động:** cộng dồn `ActiveLP` từ tháng thăng tiến đến tháng tính: <50 → 100%; ≥50 → 105%; ≥100 → 115%.
- **Hệ số ngày làm việc:** như LP toàn thời gian (tháng làm việc thứ nhất tính từ ngày bổ nhiệm).
- **Bảo đảm 6 tháng đầu:** PM gia nhập đến hết Q2/2027 được FA tối thiểu 0,5 triệu × hệ số ngày làm việc trong 6 tháng đầu kể từ tháng thăng tiến/bổ nhiệm.
- **Bảo đảm theo quý** (xét ở cuối quý trước, bù vào kỳ 4 tháng thứ 3 quý sau, cấn trừ FA đã chi):
  - PM ≤ 12 tháng kể từ thăng tiến và quý xét ≥ 8 lượt LP hoạt động → 0,5 triệu/ngày;
  - PM khác, quý xét ≥ 6 lượt LP hoạt động → 0,15 triệu/ngày.
- **Không chi:** chức danh không phải PM/PD, HĐ đại lý chấm dứt, tình trạng SE cuối tháng.

Mọi mức và ngưỡng nằm trong **Bảng tham số chính sách**.

## 5. Giả định cần anh/chị xác nhận

| # | Điểm chưa rõ trong văn bản 006 | Cách tool đang xử lý |
|---|---|---|
| 1 | Điều kiện chấm công tối thiểu (≥70% hoặc ≥15 ngày như LP) có áp dụng cho PM không | **Không áp dụng** (tham số `MinAttendancePct`, `MinWorkDays` = 0). Muốn áp dụng thì sửa tham số |
| 2 | "Gia nhập" trong bảo đảm 6 tháng đầu tính từ ngày nào | Ngày thăng tiến PM (`PROMOTION_DATE`); trống thì ngày bổ nhiệm |
| 3 | Bảo đảm 6 tháng đầu trả **hàng tháng** hay quyết toán theo quý | Trả hàng tháng, như mức sàn của FA tháng |
| 4 | Bảo đảm 6 tháng đầu có cần điều kiện doanh số không | **Không**: văn bản không nêu điều kiện |
| 5 | PM có cần là toàn thời gian (FT) không | Không xét FT/PT |
| 6 | Tháng thứ 1 của PM mới thăng tiến từ LP: hệ số ngày tính theo ngày bổ nhiệm hay ngày thăng tiến | Theo ngày bổ nhiệm (hợp đồng đại lý), như định nghĩa "tháng làm việc" trong 005 |

## 6. Kết quả kỳ vọng tháng 10/2026 với dữ liệu thật

8 PM bổ nhiệm 03/10/2026, chưa có `FAPM_Input` (FYC nhóm = 0) → phụ cấp theo công thức = 0, nhưng **được bảo đảm 6 tháng đầu**:

| PM | Số công | Hệ số ngày | FA |
|---|---|---|---|
| 7 PM đi làm 20 ngày | 20 | 20,0 | **10.000.000** |
| 00000010 | 18 | 18,0 | **9.000.000** |
| Tổng | | | **79.000.000** |

Chi tiết trong `ket-qua/FAPM_202610_Ky1_thuc.xlsx` (không đưa lên repo). Số này phụ thuộc giả định 2, 3, 4 ở mục 5.
