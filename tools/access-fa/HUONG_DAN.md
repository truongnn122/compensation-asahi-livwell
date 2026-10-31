# Tool tính Phụ cấp cố định (FA) – Microsoft Access

Căn cứ: Thông báo **ALL/AD/2210-005** (22/10/2026), Phụ lục 1 mục **A.II**, các định nghĩa D.k, D.l, D.s, D.z.

## 1. Bộ file

| File | Vai trò |
|---|---|
| `Build_FA_Tool.vbs` | Script dựng tool, chạy 1 lần |
| `modFA.bas` | Module tính FA (VBA). Đây là file duy nhất thay mới khi nâng cấp tool |
| `modUpdate.bas` | Nút **Cập nhật tool** và **Sao lưu dữ liệu**. Cài 1 lần, hầu như không đổi |
| `FA_Input_Mau.xlsx` | File mẫu nhập liệu: 10 LP test và sheet kết quả kỳ vọng để đối soát |
| `kiem-tra/fa_reference.py` | Bản tính đối chiếu bằng Python, dùng cùng logic với tool |

## 2. Cài đặt (máy Windows có Microsoft Access 2010 trở lên)

1. Chép `Build_FA_Tool.vbs`, `modFA.bas` và `modUpdate.bas` vào cùng một thư mục.
2. Double-click `Build_FA_Tool.vbs`. Chờ khoảng 30 giây đến khi hiện thông báo "Hoàn tất".
3. Script tạo ra 2 file:
   - `FA_Data.accdb`: chứa dữ liệu, tham số và kết quả. Nếu nhiều người dùng chung, đặt file này ở thư mục mạng.
   - `FA_App.accdb`: giao diện. Mỗi người giữ 1 bản trên máy mình.
4. Mở `FA_App.accdb` và bấm **Enable Content**. Nên thêm thư mục vào *File > Options > Trust Center > Trusted Locations* để lần sau không phải bấm lại.
5. Nếu đã chuyển `FA_Data.accdb` sang thư mục khác: bấm **Kết nối file dữ liệu** rồi chọn lại file.

## 3. Cập nhật tool và sao lưu

**Nâng cấp tool (không mất dữ liệu):**
1. Tải file `modFA.bas` mới và chép vào cùng thư mục với `FA_App.accdb`.
2. Mở tool, bấm **Cập nhật tool**.
3. Tool tự sao lưu dữ liệu, nạp code mới, rồi báo số phiên bản.

Code nằm trong `FA_App.accdb`, dữ liệu nằm trong `FA_Data.accdb`, nên cập nhật code không động tới dữ liệu.

**Tự sao lưu:** trước mỗi lần *nhập Excel*, *chốt kỳ* và *cập nhật tool*, tool chép `FA_Data.accdb` vào thư mục `Backup` nằm cạnh file dữ liệu. Tên bản sao có ngày giờ, ví dụ `FA_Data_20261031_143000.accdb`. Muốn lấy lại một bản: đóng tool, chép bản backup đè lên `FA_Data.accdb`.

**Bản đã cài trước ngày 31/10/2026 (cài thêm 1 lần):**
1. Alt+F11 → **File → Import File…** → chọn `modUpdate.bas`.
2. Đặt con trỏ chuột vào dòng `Sub CaiDatCapNhat`, bấm **F5**. Form chính sẽ có thêm 2 nút.
3. Bấm **Cập nhật tool** để nạp `modFA.bas` mới.

## 4. Danh sách Agent (nạp hàng tháng)

- Bấm **Danh sách Agent** trên form chính. Bảng hiện đủ 98 cột theo đúng thứ tự file xuất từ hệ thống, có ô tìm theo mã hoặc họ tên, lọc theo chức danh, nút **Nạp file Excel** và **Xuất Excel**.
- Mỗi tháng nạp file xuất hệ thống, chỉ cần sheet tên `Agent` và có cột `Period`:
  - `tblAgent` được cập nhật thành thông tin **mới nhất**: agent mới được thêm vào, agent đã có được cập nhật.
  - `tblAgentHist` lưu **bản chụp riêng của từng tháng** theo `Period`. Nạp lại cùng một tháng thì chỉ ghi đè tháng đó.
- Tính FA tháng nào thì tool dùng chức danh, ngày bổ nhiệm và FT/PT **của đúng tháng đó** trong `tblAgentHist`. Nếu chưa nạp file của tháng đó thì dùng bản mới nhất.
- Cột trùng tên trong file được đặt tên riêng:
  - `INDIRECT 2_NAME` lần thứ nhất → `INDIRECT 2_CODE`;
  - `CHANNEL` → `CHANNEL (2)`;
  - `CONTRACT_TYPE` → `CONTRACT_TYPE Agent` / `CONTRACT_TYPE Leader`.
- File có thông tin cá nhân (CCCD, tài khoản ngân hàng): nên đặt mật khẩu cho `FA_Data.accdb` (*File → Encrypt with Password*) và giới hạn quyền truy cập thư mục.

## 5. Quy trình mỗi kỳ

1. **Nhập dữ liệu từ Excel** theo đúng mẫu `FA_Input_Mau.xlsx` (sheet `Agent`, `Holiday`, `FA_Input`).
   - Sheet `Agent` nhận được **cả file xuất từ hệ thống**: chỉ cần dán nguyên file xuất vào sheet `Agent`. Tool tự tìm các cột `AGENT_CODE`, `FULL_NAME`, `APPOINTMENT_DATE`, `AGENCY_TYPE` (FT/PT), `TERMINATION_DATE`, `TITLE (Code)`, `AGENT_STATUS`, còn các cột khác được bỏ qua và không lưu vào tool.
2. Nhập tháng (YYYYMM) và chọn kỳ chi trả (1–4), sau đó bấm **Chạy tính FA**.
3. **Xem kết quả** hoặc **Xuất bảng kê Excel**. Cột *Chi kỳ này* là số tiền cần chi.
4. Sau khi đã đối soát và chi trả, bấm **Chốt kỳ**. Kỳ đã chốt không thể tính lại hay bị nhập đè.

**Chia chi trả theo 4 kỳ (mục D.l):** mỗi kỳ nhập số **lũy kế** (FYC, số hợp đồng) tính đến ngày chốt ACK của kỳ đó. Tool tính lại FA và chỉ chi phần chênh lệch so với các kỳ trước. Phần chênh lệch có thể ra số âm, nghĩa là cần thu hồi.

**Dùng chung nhiều người:**
- Khi một người đang chạy tính một tháng/kỳ, người khác bị chặn trong 30 phút.
- Không cho tính lại kỳ trước nếu kỳ sau của cùng tháng đã có kết quả.

## 6. Công thức đã cài

`FA = Phụ cấp/ngày × Hệ số ngày làm việc × %Hệ số K2 × %Hệ số MDRT × %Hệ số HĐ`

- **Phụ cấp/ngày:** lấy mức cao nhất mà LP thỏa, giữa Lựa chọn 1 và Lựa chọn 2. Các mức là 0,2 / 0,3 / 0,5 / 0,7 triệu.
- **%attendance** = (số ngày đi làm thực tế + 1) ÷ số ngày làm việc của tháng, tối đa 100%.
- **Hệ số ngày làm việc** = %attendance × 22.
- **Tháng làm việc thứ nhất:**
  - mẫu số %attendance = số ngày làm việc tính từ ngày bổ nhiệm;
  - hệ số ngày làm việc = số ngày đi làm thực tế, tối đa 22.
- **Điều kiện nhận FA:** đồng thời thỏa các điều kiện sau:
  - chức danh là LP (PM, PD... không được tính FA);
  - là LP toàn thời gian;
  - đã hoàn tất huấn luyện bắt buộc;
  - %attendance ≥ 70% hoặc đi làm ≥ 15 ngày;
  - cuối tháng không ở tình trạng SE;
  - hợp đồng đại lý chưa chấm dứt.
- **Bảo đảm LP mới:** trong 2 tháng làm việc đầu, FA tối thiểu = 0,3 triệu × hệ số ngày làm việc, khi LP có lũy kế ≥ 0,5 hợp đồng.
- **Bảo đảm theo quý:** áp dụng cho LP có ≥ 6 hợp đồng và ≥ 24 triệu FYC phát sinh trong quý trước.
  - Mức bảo đảm = 0,15 triệu × hệ số ngày làm việc của từng tháng trong quý tiếp theo.
  - Bù 1 lần vào **kỳ 4 của tháng thứ 3**, sau khi trừ FA đã chi trong quý.
- **Làm tròn (mục D.s):** tỷ lệ % làm tròn half-up đến 1 chữ số thập phân; tiền làm tròn đến đồng.

Mọi mức tiền, tỷ lệ và ngưỡng nằm trong **Bảng tham số chính sách**. Khi chính sách thay đổi, chỉ cần sửa bảng, không phải sửa code.

## 7. Giả định cần anh/chị xác nhận

| # | Điểm văn bản chưa rõ | Cách tool đang xử lý |
|---|---|---|
| 1 | Hệ số ngày làm việc có làm tròn không | Làm tròn 1 chữ số thập phân. Ví dụ trang 3: 95,5% × 22 = 21,01 → **21,0** |
| 2 | Tháng 1 chưa đủ 0,5 HĐ, sang tháng 2 mới đạt | Tháng 2 **bù** phần bảo đảm của tháng 1 (cột *Bù bảo đảm tháng 1*) |
| 3 | Mức bảo đảm (LP mới, theo quý) có nhân hệ số K2/MDRT/HĐ không | **Không**: chỉ tính mức/ngày × hệ số ngày làm việc |
| 4 | Điều kiện tối thiểu có áp dụng cho các khoản bảo đảm không | **Có**: tháng không đủ điều kiện thì không được bảo đảm |
| 5 | "FA đã chi trả" dùng để cấn trừ bảo đảm quý | Gồm FA tháng, không gồm khoản bù tháng 1 |
| 6 | Xét bảo đảm quý cho Q4/2026 khi chính sách mới hiệu lực từ 01/10/2026 | Tool tự xét theo dữ liệu Q3 nếu đã nhập; không nhập thì không có bảo đảm |
| 7 | Ngưỡng 50.000đ cộng dồn (mục D.aa) áp trên tổng thu nhập | **Chưa áp dụng** vì tool chỉ tính FA |
| 8 | PM/PD có được nhận FA không | **Không**: mục A.II chỉ nêu "LP toàn thời gian". Danh sách chức danh được nhận FA nằm ở hằng số `FA_TITLES` trong `modFA.bas` |

## 8. Kiểm tra

Kết quả tính bằng file mẫu, tháng 10/2026, kỳ 1, phải khớp sheet `KetQua_KyVong`, cụ thể:

| LP | Kịch bản | FA |
|---|---|---|
| LP001 | Ví dụ trang 3 của văn bản | 10.500.000 |
| LP002 | Mức 0,7 triệu, K2 120%, MDRT 120% | 22.176.000 |
| LP006 | LP mới, được bảo đảm | 4.200.000 |
| LP004, LP005, LP008, LP009, PM001 | Các trường hợp không đủ điều kiện | 0 |
