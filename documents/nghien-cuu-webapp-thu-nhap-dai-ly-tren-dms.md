# Nghiên cứu: Webapp tính thu nhập Đại lý bảo hiểm nhân thọ trên nền DMS

Ngày lập: 23/10/2026. Phạm vi: các giải pháp tính hoa hồng/thưởng cho đại lý (Compensation / Commission module) chạy trên hoặc tích hợp với DMS (Distribution Management System), gồm nhà cung cấp quốc tế, ứng dụng đại lý tại Việt Nam, khung pháp lý và gợi ý kiến trúc cho dự án Asahi Livwell.

---

## 1. Kết luận nhanh

1. **Không có sản phẩm "đóng gói" nào cho thị trường Việt Nam.** Các DMS quốc tế (Majesco, Sircon/Vertafore, Equisoft, EIS, Sapiens) đều có module compensation, nhưng giá cao, triển khai dài và không có sẵn quy tắc thu nhập theo kiểu Việt Nam (FYC, K2/K13, override theo cấp bậc, thưởng tuyển dụng).
2. **Các hãng tại Việt Nam tự xây (in-house).** Prudential (PRUForce), Manulife (M-PA / MaxX D) đưa tính năng xem thu nhập, ước tính hoa hồng vào app đại lý. Đây là chuẩn kỳ vọng của tư vấn viên hiện nay.
3. **Mẫu kiến trúc đáng học nhất:** engine tính toán tách rời khỏi UI/DB, tính từ **nhật ký sự kiện bất biến** (event log) để hồi tố (clawback, truy thu, tính lại) sạch sẽ, và mọi đồng tiền truy vết được về đúng "quy tắc + sự kiện" sinh ra nó. Dự án hiện tại (Next.js + TypeScript + Firebase) hoàn toàn phù hợp để xây theo mẫu này.
4. **Ràng buộc pháp lý bắt buộc cài vào engine:** trần hoa hồng theo nghiệp vụ (Điều 51 TT 67/2023/TT-BTC), trần thưởng/hỗ trợ đại lý (Điều 52), hoa hồng tính trên **phí thực thu**, và từ 01/01/2026 đại lý cá nhân không được làm cho 2 công ty nhân thọ cùng lúc.

---

## 2. Bức tranh giải pháp

### 2.1 DMS quốc tế có module compensation

| Giải pháp | Phân khúc | Năng lực compensation nổi bật (theo tài liệu nhà cung cấp) | Nhận xét cho Việt Nam |
|---|---|---|---|
| **Majesco Distribution Management / Life DistributionPlus** | Carrier lớn, Life & Annuity + P&C | Rules-based commission engine; bảng tỷ lệ; hệ thống phân cấp n-tier; tạm ứng hoa hồng; **charge-back theo sự kiện hợp đồng**; chia sẻ/điều chỉnh hoa hồng; contingent commission, bonus, profit sharing; cấu hình không cần sửa code | Mạnh và toàn diện, nhưng nặng. Phù hợp làm **chuẩn tính năng** để đối chiếu |
| **Sircon Compensation (Vertafore)** | Carrier Mỹ | N-level hierarchy, MGA/FMO theo sản phẩm/kênh, tự động re-code; chu kỳ thanh toán gồm earnings + advances + held payments + adjustments; bonus đa tầng có workflow duyệt; **cổng tự phục vụ** cho đại lý xem chi tiết hoa hồng/thưởng; nhà cung cấp nêu giảm ~50% thời gian xử lý | Đáng học ở **chu kỳ thanh toán** và **self-service statement** |
| **Equisoft Centralize** | Nhà phân phối / MGA (Canada) | Commission engine đa cấp; lịch hoa hồng, **override rate, sharing %**; đối soát hoa hồng kỳ vọng vs nhận được; loans, caps, chargebacks, bonus; chuyển block of business; 50+ báo cáo; cổng đại lý | Đáng học ở **đối soát** và **override/sharing** |
| **EIS Compensation Manager / Producer Management** | Carrier | Cấu trúc compensation cho producer, broker portal xem chi tiết | Ít tài liệu công khai |
| **Sapiens DigitalSuite (commission processing)** | Carrier vừa và lớn | Xử lý hoa hồng tích hợp với PAS của Sapiens | Chỉ hợp lý nếu đã dùng PAS Sapiens |
| **ICM thuần (Varicent, Xactly, CaptivateIQ, Performio)** | Sales comp đa ngành | Plan builder no-code, tiers/accelerators, dispute workflow, GL export. Varicent có khách hàng ngành bảo hiểm | Linh hoạt nhưng thiếu khái niệm bảo hiểm (K2, lapse, chargeback theo hợp đồng) phải tự mô hình hóa |

Nguồn tổng hợp (Finantrix buyer guide) xếp Sapiens và EXL ComPAS là "Leader" về commission calculation cho carrier; các danh sách "top 10" trên GitNux/Zipdo không có phương pháp chấm điểm rõ nên chỉ tham khảo.

### 2.2 Ứng dụng đại lý tại Việt Nam (chuẩn kỳ vọng của tư vấn viên)

| Hãng | Ứng dụng | Tính năng liên quan thu nhập (theo mô tả công khai) |
|---|---|---|
| Prudential | **PRUForce** | Tra cứu khách hàng, hợp đồng, tình hình kinh doanh; **công cụ ước tính hoa hồng dựa trên phí thực thu** |
| Manulife | **M-PA (Manulife Advisor)**, MaxX D | Theo dõi FYP, FYC, APE theo tháng; **bảng chi tiết thu nhập**, chính sách; funnel bán hàng và tuyển dụng; phân tích theo hệ thống (hierarchy) và hiệu suất |
| Dai-ichi, AIA, FWD, Generali, Sun Life | App đại lý nội bộ | Không có mô tả công khai về màn hình thu nhập; FWD Việt Nam chủ yếu qua banca (Vietcombank, Agribank, HDBank) |
| Vymo | Nền tảng sales engagement | Nhiều khách hàng khu vực (AIA, Generali, Sun Life), **không xác nhận** được triển khai tại Việt Nam |

Không tìm thấy nhà cung cấp Việt Nam (FPT IS, TMA, CMC...) công bố sản phẩm compensation cho nhân thọ. Kết luận: **thị trường trong nước đang tự xây**, và sản phẩm của Asahi Livwell cần ít nhất ngang chuẩn PRUForce / M-PA ở phần "xem thu nhập + ước tính hoa hồng".

### 2.3 Mã nguồn mở tham khảo

- **OpenICM** (github.com/kyssta-exe/OpenICM): Next.js 15 App Router, TypeScript, PostgreSQL, Prisma. Điểm mạnh: engine thuần TypeScript trong `src/engine/` không phụ thuộc DB, tính hoa hồng là **truy vấn xác định trên event log bất biến**; hỗ trợ tiers, accelerators, clawbacks, splits, sửa deal hồi tố; audit/approval; xuất GL/payroll. Dự án còn rất mới (ít commit), chỉ nên **mượn thiết kế**, không nên dùng trực tiếp.
- Odoo commission modules, Lovable CommCalc: chỉ là add-on/template, không đủ cho hệ thống phân cấp bảo hiểm.

---

## 3. Khung pháp lý phải cài vào engine (Việt Nam)

Nguồn: Thông tư 67/2023/TT-BTC (hiệu lực 02/11/2023, một số tỷ lệ áp dụng cho hợp đồng phát hành từ 01/7/2024) và Luật số 139/2025/QH15. Các con số dưới đây lấy từ trang pháp luật thứ cấp, **cần đối chiếu văn bản gốc trước khi cấu hình**.

| Nội dung | Quy định | Hàm ý thiết kế |
|---|---|---|
| Cơ sở tính hoa hồng | Tính trên **phí bảo hiểm thực tế thu được** của từng hợp đồng | Engine phải chạy theo sự kiện **thu phí**, không theo phí ghi nhận/issue |
| Trần hoa hồng (Điều 51) | Theo nghiệp vụ và năm hợp đồng. Liên kết chung/đơn vị: nguồn thứ cấp nêu 30% năm 1, 20% năm 2, 10% năm sau (từ 01/7/2024, giảm so với 40% trước đó); hợp đồng ≤ 1 năm: 20%; hưu trí: 3% tổng phí; nhân thọ nhóm: 50% mức cá nhân | Bảng tỷ lệ phải **versioned theo ngày phát hành hợp đồng**; validator chặn vượt trần |
| Trần thưởng/hỗ trợ (Điều 52) | Tổng thưởng, hỗ trợ, quyền lợi khác mỗi năm tài chính ≤ 20% phí thực thu của HĐ ≤ 1 năm **+ 30% phí khai thác năm đầu thực thu** của HĐ > 1 năm | Cần báo cáo **ngân sách thưởng theo năm tài chính** và cảnh báo khi gần trần |
| Độc quyền đại lý | Từ 01/01/2026, cá nhân đang làm đại lý cho 1 công ty nhân thọ không được làm cho công ty nhân thọ khác | Ảnh hưởng chính sách tuyển dụng/chuyển hệ thống; cần lưu lịch sử mã đại lý và kiểm tra khi onboarding |
| Thuế TNCN | Thu nhập đại lý thuộc nhóm chịu thuế; khấu trừ tại nguồn | Bảng kê phải tách **thu nhập gộp, thuế khấu trừ, thực nhận** |

---

## 4. Tính năng chuẩn của một module tính thu nhập trên DMS

Tổng hợp từ Majesco, Sircon, Equisoft, OpenICM và thực tế vận hành:

**A. Dữ liệu đầu vào (từ DMS / PAS)**
- Hồ sơ đại lý: mã, cấp bậc, ngày hiệu lực cấp bậc, trạng thái (active/terminated), hợp đồng đại lý, **lịch sử thay đổi**.
- Cây phân cấp (hierarchy) **có hiệu lực theo thời gian**: ai quản lý ai, tại thời điểm nào; chuyển hệ thống (re-code), chuyển block hợp đồng.
- Sự kiện hợp đồng: issue, thu phí (FYP/RYP), hoàn phí, free-look hủy, lapse, reinstatement, surrender, giảm phí, thay đổi sản phẩm.
- Danh mục sản phẩm: tỷ lệ hoa hồng theo sản phẩm × năm hợp đồng × kênh, hệ số APE.

**B. Engine tính toán**
- Hoa hồng năm đầu (FYC), hoa hồng tái tục, override cho các cấp quản lý trực tiếp/gián tiếp theo hierarchy tại thời điểm thu phí.
- Thưởng: tháng/quý/năm theo FYC, APE, số case, tỷ lệ duy trì (K2/K13), thưởng tuyển dụng, thưởng thăng cấp, contest.
- **Điều chỉnh hồi tố**: chargeback/clawback khi lapse/hủy, truy thu, tạm ứng và cấn trừ, giữ lại (hold) khi vi phạm.
- Tiers, sliding scale, accelerators; cap theo trần pháp lý.
- Tính lại (re-run) một kỳ đã khóa khi có sự kiện muộn, tạo **bút toán chênh lệch** cho kỳ sau thay vì sửa kỳ cũ.

**C. Quy trình thanh toán**
- Chu kỳ: tính thử (dry-run), đối soát, duyệt theo phân quyền, khóa kỳ, xuất file ngân hàng, bút toán kế toán, khấu trừ thuế.
- Đối soát: kỳ vọng vs thực chi; chênh lệch được gắn cờ.

**D. Trải nghiệm đại lý (self-service)**
- Dashboard thu nhập: tháng này, dự kiến, so sánh kỳ trước; FYP/FYC/APE; tiến độ tới mốc thưởng.
- **Bảng kê drill-down**: từ tổng → loại thu nhập → từng hợp đồng → quy tắc áp dụng. Đây là điểm yếu phổ biến (PDF hằng tháng, không drill-down) mà Sircon và Equisoft nhấn mạnh.
- Công cụ **ước tính thu nhập "what-if"** (như PRUForce): nhập phí/sản phẩm → hoa hồng dự kiến.
- Khiếu nại/tra soát (dispute) có luồng xử lý và SLA.

**E. Quản trị và tuân thủ**
- Plan builder có **phiên bản và ngày hiệu lực**; mô phỏng tác động chính sách mới trên dữ liệu thật trước khi ban hành.
- Audit trail: ai đổi gì, khi nào; mọi dòng tiền truy vết về quy tắc + sự kiện.
- Báo cáo ngân sách thưởng theo trần Điều 52; báo cáo cho cơ quan quản lý.

---

## 5. Gợi ý kiến trúc cho dự án Asahi Livwell

Repo hiện tại đã có: Next.js 16 App Router, TypeScript strict, Firebase/Firestore, exceljs, TanStack Table, module tuyển dụng (recruitments) và người dùng/phân quyền. Đề xuất phát triển module thu nhập theo hướng:

1. **Tách engine thuần TypeScript** (`src/lib/compensation/engine/`): input là danh sách sự kiện + cấu hình chính sách, output là các dòng earning. Không gọi Firestore bên trong engine để dễ unit test (có skill `jest` sẵn trong repo).
2. **Event log bất biến** trong Firestore: `premium_collected`, `policy_lapsed`, `agent_promoted`, `hierarchy_changed`... Mỗi kỳ tính là một snapshot; sự kiện muộn sinh điều chỉnh ở kỳ sau.
3. **Policy-as-config có version**: bảng tỷ lệ, công thức thưởng, trần pháp lý lưu dạng dữ liệu với `effectiveFrom/effectiveTo`, không hard-code.
4. **Import từ DMS/PAS bằng Excel** trước (exceljs đã có), API sau. Người dùng đang làm form mẫu Excel nên luồng Excel → validate → tính thử → duyệt → xuất bảng kê là bước đi thực tế nhất.
5. **Ba màn hình ưu tiên**: (a) Cấu hình chính sách và mô phỏng; (b) Chạy kỳ thanh toán và duyệt; (c) Bảng kê thu nhập đại lý có drill-down và ước tính what-if.
6. **Kiểm soát tuân thủ ngay trong engine**: validator trần Điều 51/52, cảnh báo khi vượt, log lý do override thủ công.

### Lộ trình gợi ý
| Giai đoạn | Nội dung | Kết quả |
|---|---|---|
| 1 | Mô hình dữ liệu + engine FYC/renewal/override + import Excel + bảng kê cơ bản | Thay thế tính tay bằng Excel |
| 2 | Thưởng (tháng/quý/năm, K2, tuyển dụng), chargeback/clawback, khóa kỳ, duyệt, xuất ngân hàng | Vận hành thanh toán end-to-end |
| 3 | Portal đại lý: dashboard, drill-down, what-if, tra soát | Ngang chuẩn PRUForce / M-PA |
| 4 | Mô phỏng chính sách, so sánh đối thủ, báo cáo trần pháp lý | Hỗ trợ thiết kế chính sách |

---

## 6. Điểm cần bạn xác nhận

1. **DMS hiện tại của công ty là hệ thống nào** (tự xây, Majesco, Sapiens, hay khác)? Điều này quyết định cách lấy dữ liệu hợp đồng/thu phí (API hay file).
2. Thu nhập tính theo **phí thực thu theo ngày** hay chốt theo tháng? Ảnh hưởng trực tiếp tới thiết kế event log.
3. Cây phân cấp có cho phép **đổi quản lý giữa kỳ** không, và override tính theo hierarchy tại ngày thu phí hay ngày cuối kỳ?

---

## Nguồn tham khảo

- Majesco Distribution Management: https://www.majesco.com/distribution-software-insurance-solutions/distribution-management/ ; brochure https://www.majesco.com/wp-content/uploads/2019/03/Majesco-Distribution-Management-Final.pdf ; Life DistributionPlus https://www.majesco.com/wp-content/uploads/2019/03/Majesco-Life-DistributionPlus.pdf ; Celent profile https://www.celent.com/solutions/513032081
- Sircon Compensation (Vertafore): https://vertafore.com/products/insurance-commission-tracking-software/sircon-compensation ; datasheet https://www.vertafore.com/sites/default/files/documents/Sircon-Compensation-Datasheet-2023.pdf
- Equisoft Centralize: https://www.equisoft.com/products/agency-management-system-centralize ; Celent https://www.celent.com/solutions/232592064 ; bài về agent portal https://www.equisoft.com/insights/insurance/insurance-agent-client-portals-all-you-need-to-know
- EIS Producer Management: https://www.eisgroup.com/producer-management/
- Finantrix buyer guide (commission calculation for carriers): https://www.finantrix.com/buyer-guides/commission-calculation-software-insurance-carriers
- Provena guide DMS 2026: https://www.provena-ai.com/blog/insurance-distribution-management-software-guide
- Kennect, insurance commissions at scale: https://www.kennect.io/blog/insurance-commissions
- AgentTech, thiết kế commission grid: https://agenttech.io/blog/agency-commission-grid-design
- OpenICM (mã nguồn mở): https://github.com/kyssta-exe/OpenICM
- G2 so sánh Performio / Varicent / CaptivateIQ: https://www.g2.com/compare/performio-vs-varicent-varicent-vs-captivateiq
- PRUForce (App Store): https://apps.apple.com/vn/app/pruforce/id1619489131
- Manulife Advisor M-PA (Google Play): https://play.google.com/store/apps/details?id=com.manulife.vn.dist.mpa ; MaxX D https://vir.com.vn/manulife-vietnam-launches-innovative-new-personal-virtual-assistant-for-agents-86005.html
- Vymo OnboardIQ: https://vymo.com/us/onboardiq/
- Thông tư 67/2023/TT-BTC toàn văn: https://lawplayer.com/vn/act/vn-vbpl-163439 ; https://trolyluat.vn/van-ban/thong-tu-672023tt-btc-33872
- Trần hoa hồng nhân thọ (Điều 51): https://thuvienphapluat.vn/phap-luat/ho-tro-phap-luat/ty-le-hoa-hong-dai-ly-toi-da-doi-voi-hop-dong-bao-hiem-nhan-tho-theo-thong-tu-672023ttbtc-ra-sao-126991.html ; https://gonnapass.com/ty-le-hoa-hong-dai-ly-bao-hiem/ ; https://xaydungchinhsach.chinhphu.vn/quy-dinh-moi-ve-muc-hoa-hong-dai-ly-bao-hiem-119231112081311873.htm
- Trần thưởng/hỗ trợ đại lý (Điều 52): https://luatvietnam.vn/tin-van-ban-moi/gioi-han-muc-thuong-cho-dai-ly-bao-hiem-186-96183-article.html ; https://luatvietnam.vn/bao-hiem/diem-moi-cua-thong-tu-67-2023-tt-btc-ve-kinh-doanh-bao-hiem-563-96146-article.html
- Giảm hoa hồng năm đầu từ 01/7/2024: https://www.tinnhanhchungkhoan.vn/lan-dau-tien-giam-hoa-hong-bao-hiem-nam-dau-post336070.html ; https://www.tinnhanhchungkhoan.vn/dai-ly-bao-hiem-ban-khoan-thu-nhap-post346932.html
- Luật 139/2025/QH15, đại lý không làm 2 công ty nhân thọ từ 2026: https://znews.vn/tu-2026-tu-van-vien-bao-hiem-nhan-tho-het-duoc-chay-2-chan-post1609935.html
- Thuế TNCN thu nhập đại lý từ 2026: https://thuvienphapluat.vn/phap-luat/tu-nam-2026-thu-nhap-tu-kinh-doanh-dai-ly-bao-hiem-cua-ca-nhan-cu-tru-chiu-thue-tncn-bao-nhieu-phan-250513.html
