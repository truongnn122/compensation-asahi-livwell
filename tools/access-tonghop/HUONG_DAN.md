# Tool Tổng hợp thu nhập – Microsoft Access

Gom kết quả từ các file Access khoản thưởng (FA, thưởng tháng, thưởng quý…) thành tổng thu nhập của từng agent theo tháng/kỳ.

## 1. Vị trí trong mô hình

| File | Vai trò |
|---|---|
| `Data.accdb` | Dữ liệu gốc dùng chung (Tbl_Agent, Tbl_Attendance…) |
| `FA_App.accdb`, `ThuongThang.accdb`, … | Mỗi khoản thưởng 1 file: tính, chốt kỳ, lưu bảng kết quả |
| **`TongHop.accdb`** (file này) | Chỉ **đọc** bảng kết quả của từng khoản thưởng, cộng lại theo agent, lưu vào `tblTongHop` |

Tool Tổng hợp không ghi gì vào file khoản thưởng. Mỗi khoản thưởng chỉ là **1 dòng khai báo** trong bảng `tblBonusSource`, thêm khoản mới không cần sửa code.

## 2. Cài đặt (1 lần)

1. Mở Access → **Blank database** → đặt tên `TongHop.accdb`, lưu cùng thư mục với các file khoản thưởng.
2. Bấm **Alt+F11** → **File → Import File…** → chọn `modTongHop.bas`.
3. Bấm **Ctrl+G**, gõ `THSetup` rồi **Enter**. Lệnh này tạo 3 bảng và form chính.
4. Đóng hẳn file rồi mở lại. Form **Tổng hợp** tự mở.
5. Bấm **Thêm file khoản thưởng** → chọn `FA_App.accdb`. Tool nhận ra file FA và tự điền cấu hình.

## 3. Dùng hàng kỳ

1. Nhập **Tháng** (YYYYMM) và **Kỳ chi trả** (1–4, hoặc `0` để gom cả tháng).
2. Bấm **1. Tổng hợp**. Tool đọc từng file khoản thưởng, ghi vào `tblTongHop` và ghi 1 dòng nhật ký cho mỗi khoản (số agent, tổng tiền, kỳ đã chốt chưa).
3. **2. Xem bảng tổng hợp:** mỗi agent 1 dòng, mỗi khoản thưởng 1 cột, cột cuối là *Tổng thu nhập*.
4. **3. Xem chi tiết từng khoản** và **4. Xuất Excel** khi cần.

**Ô "Chỉ lấy kỳ đã CHỐT"** (mặc định bật): khoản thưởng nào chưa chốt kỳ đó thì bị bỏ qua và ghi nhật ký *"kỳ chưa chốt - bỏ qua"*. Bỏ tick để xem số tạm; dòng tổng hợp sẽ được đánh dấu `PeriodLocked = No`.

Tổng hợp lại cùng tháng/kỳ sẽ **ghi đè** kết quả tổng hợp cũ của tháng/kỳ đó. Nhật ký (`tblCollectLog`) giữ lại toàn bộ lịch sử các lần tổng hợp.

## 4. Thêm khoản thưởng mới

Bấm **Thêm file khoản thưởng**, nhập mã (ví dụ `TTHANG`), rồi điền trong bảng `tblBonusSource`:

| Cột | Ý nghĩa | Ví dụ với FA |
|---|---|---|
| `ResultTable` | Bảng kết quả trong file khoản thưởng | `tblFA_Result` |
| `AgentField` | Cột mã agent | `AgentCode` |
| `NameField` | Cột họ tên (để trống nếu không có) | `AgentName` |
| `PeriodField` | Cột tháng (YYYYMM) | `Period` |
| `KyField` | Cột kỳ chi trả (để trống nếu khoản đó không chia kỳ) | `Ky` |
| `AmountField` | Cột số tiền chi kỳ này | `Payable` |
| `StatusTable` | Bảng trạng thái chốt kỳ, có cột `Status` = `LOCKED` (để trống thì coi như luôn đã chốt) | `tblPeriodStatus` |
| `SortOrder`, `Active` | Thứ tự cột và có lấy hay không | |

Các file khoản thưởng tôi dựng sau này sẽ dùng cùng tên cột với FA, nên chỉ cần chọn file là xong.

## 5. Bảng dữ liệu

| Bảng | Nội dung |
|---|---|
| `tblBonusSource` | Khai báo các file khoản thưởng |
| `tblTongHop` | Kết quả tổng hợp: tháng, kỳ, agent, khoản thưởng, số tiền, kỳ đã chốt chưa, file nguồn, thời điểm và người tổng hợp |
| `tblCollectLog` | Nhật ký từng lần tổng hợp |

## 6. Chưa làm (chờ chính sách)

- Khấu trừ thuế TNCN trên tổng thu nhập.
- Thu nhập dưới 50.000đ cộng dồn sang kỳ sau (mục D.aa).
- Bảng kê chi trả theo mẫu kế toán.
