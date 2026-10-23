# Khung module tính thu nhập Đại lý (bản giả định để xem)

Mục đích: nhìn toàn cảnh module trước khi làm thật. **Mọi tỷ lệ, cấp bậc, ngưỡng trong tài liệu này là giả định** để minh họa. Khi có quy chế thật, chỉ thay số trong phần 1, cấu trúc các phần sau giữ nguyên.

---

## 1. Chính sách mẫu (giả định)

### 1.1 Cấp bậc và cây phân cấp
| Mã | Cấp bậc | Quản lý trực tiếp |
|---|---|---|
| FC | Tư vấn viên | UM |
| UM | Trưởng nhóm | BM |
| BM | Trưởng phòng | AD |
| AD | Giám đốc kinh doanh | Công ty |

### 1.2 Chỉ tiêu dùng trong tính toán
| Chỉ tiêu | Định nghĩa |
|---|---|
| FYP | Phí năm đầu **thực thu** trong kỳ |
| APE | Phí quy năm: đóng năm ×1, nửa năm ×2, quý ×4, tháng ×12, đóng một lần ×10% |
| FYC | Hoa hồng năm đầu = FYP × tỷ lệ HH năm 1 theo sản phẩm |
| RYC | Hoa hồng tái tục = phí tái tục thực thu × tỷ lệ HH năm 2, 3+ |
| K2 | Tỷ lệ duy trì hợp đồng năm 2 (theo phí) của cá nhân hoặc nhóm |
| Case | Số hợp đồng phát hành và thu phí trong kỳ |

### 1.3 Bảng tỷ lệ hoa hồng theo sản phẩm (giả định, nằm dưới trần TT 67/2023)
| Nhóm sản phẩm | Năm 1 | Năm 2 | Năm 3+ | Trần pháp lý năm 1 (tham chiếu) |
|---|---|---|---|---|
| Liên kết chung (UL) | 28% | 15% | 5% | 30% (HĐ phát hành từ 01/7/2024) |
| Tử kỳ | 25% | 5% | 5% | theo Điều 51 |
| Hỗn hợp | 25% | 10% | 5% | theo Điều 51 |
| Bổ trợ sức khỏe (≤ 1 năm, tái tục) | 18% | 18% | 18% | 20% |

### 1.4 Override cho quản lý (tính trên FYC của hệ thống bên dưới)
| Cấp | Nhóm trực tiếp | Nhóm gián tiếp (cấp dưới của cấp dưới) |
|---|---|---|
| UM | 20% FYC của FC trong nhóm | 0 |
| BM | 10% FYC toàn phòng (gồm UM và FC) | 5% FYC phòng của BM được mình phát triển |
| AD | 5% FYC toàn khu vực | 0 |

Override tính theo **cây phân cấp tại ngày thu phí**.

### 1.5 Thưởng (giả định)
| Mã | Thưởng | Điều kiện | Mức |
|---|---|---|---|
| B-M1 | Thưởng tháng cá nhân | FYC tháng ≥ 10tr và ≥ 2 case | 10% FYC; ≥ 20tr: 15%; ≥ 40tr: 20% |
| B-Q1 | Thưởng quý cá nhân | FYC quý ≥ 60tr và K2 cá nhân ≥ 80% | 10% FYC quý |
| B-R1 | Thưởng tuyển dụng | FC mới do mình giới thiệu đạt FYC ≥ 5tr trong 3 tháng đầu | 1.000.000 đ/người |
| B-T1 | Thưởng nhóm cho UM | FYC nhóm tháng ≥ 50tr và ≥ 3 FC active | 5% FYC nhóm |
| B-P1 | Thưởng duy trì | K2 cá nhân cuối năm ≥ 85% | 3% tổng FYC năm |

### 1.6 Truy thu (chargeback)
| Sự kiện | Thời điểm | Thu hồi |
|---|---|---|
| Hủy trong thời gian cân nhắc (free-look) | bất kỳ | 100% FYC, override, thưởng đã trả liên quan |
| Hoàn phí / giảm phí | bất kỳ | theo phần phí hoàn |
| Mất hiệu lực (lapse), hủy | tháng 1 đến 12 | 100% FYC và override đã trả |
| Mất hiệu lực (lapse), hủy | tháng 13 đến 24 | 50% FYC và override đã trả |
| Khôi phục hiệu lực | bất kỳ | hoàn lại khoản đã truy thu khi thu được phí |

Truy thu được cấn trừ vào thu nhập kỳ hiện tại; phần âm chuyển sang kỳ sau (công nợ đại lý).

### 1.7 Ràng buộc pháp lý cài cứng trong engine
- Hoa hồng từng hợp đồng không vượt trần Điều 51 theo nghiệp vụ và năm hợp đồng.
- Tổng thưởng, hỗ trợ, quyền lợi khác trong năm tài chính ≤ 30% FYP thực thu (HĐ > 1 năm) + 20% phí thực thu (HĐ ≤ 1 năm) theo Điều 52.
- Thuế TNCN khấu trừ tại nguồn theo tham số cấu hình (ngưỡng và tỷ lệ cập nhật theo luật hiện hành).

---

## 2. Mô hình dữ liệu (Firestore, phác thảo)

| Collection | Vai trò | Trường chính |
|---|---|---|
| `agents` | Hồ sơ đại lý | agentCode, fullName, status, joinDate, recruiterCode, bankAccount, taxCode |
| `agentRankHistory` | Lịch sử cấp bậc | agentCode, rank, effectiveFrom, effectiveTo |
| `hierarchyEdges` | Cây phân cấp có hiệu lực theo thời gian | agentCode, managerCode, effectiveFrom, effectiveTo |
| `products` | Sản phẩm | productCode, group, termType (>1 năm / ≤1 năm) |
| `commissionSchedules` | Bảng HH có phiên bản | productCode, policyYear, rate, effectiveFrom (theo ngày phát hành HĐ) |
| `policies` | Hợp đồng | policyNo, productCode, agentCode, issueDate, paymentMode, annualPremium, status |
| `premiumEvents` | **Nhật ký sự kiện bất biến** | eventId, type, policyNo, amount, policyYear, occurredAt, importBatchId |
| `compPlans` | Chính sách thu nhập có phiên bản | planId, effectiveFrom, rules[] (tham số của mục 1) |
| `payPeriods` | Kỳ thanh toán | periodId (YYYY-MM), status (draft / calculated / approved / locked / paid) |
| `earningLines` | Dòng thu nhập đã tính | periodId, agentCode, type, amount, sourcePolicyNo, ruleId, sourceEventId |
| `adjustments` | Điều chỉnh thủ công | periodId, agentCode, amount, reason, approvedBy |
| `statements` | Bảng kê theo đại lý | periodId, agentCode, gross, chargeback, tax, net, lines[] |
| `auditLogs` | Ai đổi gì | actor, action, before, after, at |

Loại sự kiện trong `premiumEvents`: `PREMIUM_COLLECTED`, `PREMIUM_REFUNDED`, `FREE_LOOK_CANCELLED`, `POLICY_LAPSED`, `POLICY_REINSTATED`, `POLICY_SURRENDERED`, `PREMIUM_REDUCED`.

Nguyên tắc: không sửa sự kiện cũ. Sự kiện đến muộn (ví dụ lapse tháng trước mới báo) được ghi với `occurredAt` thật và tạo dòng điều chỉnh ở kỳ hiện tại.

---

## 3. Danh mục quy tắc (rule catalog)

| Rule ID | Tên | Đầu vào | Công thức | Tham số cấu hình | Có truy thu |
|---|---|---|---|---|---|
| R-FYC | Hoa hồng năm đầu | PREMIUM_COLLECTED năm 1 | amount × rate(product, year 1, issueDate) | bảng HH | Có |
| R-RYC | Hoa hồng tái tục | PREMIUM_COLLECTED năm 2+ | amount × rate(product, year n) | bảng HH | Có |
| R-OVR-UM | Override UM | FYC của FC trong nhóm tại ngày thu phí | FYC × 20% | tỷ lệ theo cấp | Có |
| R-OVR-BM | Override BM | FYC toàn phòng; FYC phòng gián tiếp | FYC × 10%; FYC × 5% | tỷ lệ theo cấp | Có |
| R-OVR-AD | Override AD | FYC toàn khu vực | FYC × 5% | tỷ lệ theo cấp | Có |
| R-B-M1 | Thưởng tháng | FYC tháng, số case | bậc thang 10/15/20% | ngưỡng, tỷ lệ | Có (theo HĐ nguồn) |
| R-B-Q1 | Thưởng quý | FYC quý, K2 | 10% nếu đạt | ngưỡng FYC, K2 | Có |
| R-B-R1 | Thưởng tuyển dụng | FYC của FC mới trong 3 tháng | 1tr khi đạt mốc | mốc, số tiền | Không |
| R-B-T1 | Thưởng nhóm UM | FYC nhóm, số FC active | 5% nếu đạt | ngưỡng | Có |
| R-CB | Truy thu | sự kiện hủy / lapse / hoàn phí | đảo ngược các dòng đã trả theo tỷ lệ thu hồi | tỷ lệ theo tháng HĐ | n/a |
| R-CAP-51 | Trần hoa hồng | mọi dòng R-FYC, R-RYC | chặn nếu rate > trần | bảng trần theo ngày phát hành | n/a |
| R-CAP-52 | Trần thưởng năm | tổng thưởng năm tài chính | cảnh báo khi ≥ 90% trần, chặn khi vượt | 30% / 20% | n/a |
| R-TAX | Thuế TNCN | thu nhập gộp lũy kế | theo tham số | ngưỡng, tỷ lệ | n/a |

Mỗi dòng thu nhập sinh ra luôn ghi `ruleId` và `sourceEventId`, nên bất kỳ số nào trên bảng kê cũng truy ngược được về sự kiện và quy tắc.

---

## 4. Luồng chạy một kỳ thanh toán

1. **Import** file Excel thu phí và trạng thái hợp đồng từ DMS → ghi vào `premiumEvents` (mỗi file là một `importBatchId`, không ghi đè).
2. **Validate**: mã đại lý tồn tại, sản phẩm có bảng HH, ngày hợp lệ, trùng lặp sự kiện.
3. **Snapshot**: dựng cây phân cấp và cấp bậc tại từng ngày thu phí trong kỳ.
4. **Tính thử (dry-run)**: engine chạy toàn bộ rule, xuất `earningLines` ở trạng thái nháp và báo cáo tổng hợp + cảnh báo trần.
5. **Đối soát**: so sánh với kỳ trước, với file Excel tính tay (giai đoạn song song), liệt kê chênh lệch lớn.
6. **Điều chỉnh thủ công** (nếu có) qua `adjustments`, bắt buộc lý do và người duyệt.
7. **Duyệt và khóa kỳ**: trạng thái `locked`, không cho tính lại; sự kiện muộn đi vào kỳ sau.
8. **Xuất**: bảng kê từng đại lý (PDF/Excel), file chi lương ngân hàng, bảng khấu trừ thuế, bút toán kế toán.

---

## 5. Khung engine (TypeScript, thuần, không gọi DB)

```ts
// src/lib/compensation/engine/types.ts
export type PremiumEventType =
  | "PREMIUM_COLLECTED" | "PREMIUM_REFUNDED" | "FREE_LOOK_CANCELLED"
  | "POLICY_LAPSED" | "POLICY_REINSTATED" | "POLICY_SURRENDERED" | "PREMIUM_REDUCED";

export interface PremiumEvent {
  eventId: string;
  type: PremiumEventType;
  policyNo: string;
  agentCode: string;        // đại lý khai thác
  productCode: string;
  policyYear: number;       // 1, 2, 3...
  amount: number;           // VND, dương
  occurredAt: string;       // ISO date
  issueDate: string;
  paymentMode: "ANNUAL" | "SEMI" | "QUARTERLY" | "MONTHLY" | "SINGLE";
}

export interface HierarchySnapshot {
  // trả về chuỗi quản lý của một đại lý tại một ngày
  managersOf(agentCode: string, atDate: string): { agentCode: string; rank: Rank }[];
  rankOf(agentCode: string, atDate: string): Rank;
}

export type Rank = "FC" | "UM" | "BM" | "AD";

export interface CompPlan {
  planId: string;
  effectiveFrom: string;
  commissionRates: Record<string, { year1: number; year2: number; year3plus: number }>;
  overrideRates: Record<Rank, { direct: number; indirect: number }>;
  monthlyBonusTiers: { minFyc: number; minCases: number; rate: number }[];
  chargebackSchedule: { fromMonth: number; toMonth: number; recoverRate: number }[];
  legalCaps: { year1ByProduct: Record<string, number>; bonusCapLongTerm: number; bonusCapShortTerm: number };
}

export interface EarningLine {
  agentCode: string;
  type: "FYC" | "RYC" | "OVERRIDE" | "BONUS" | "CHARGEBACK" | "ADJUSTMENT";
  amount: number;           // âm nếu truy thu
  ruleId: string;
  sourceEventId: string;
  sourcePolicyNo: string;
  periodId: string;
}

export interface RunInput {
  periodId: string;
  events: PremiumEvent[];           // sự kiện trong kỳ + sự kiện muộn chưa xử lý
  paidHistory: EarningLine[];       // các dòng đã trả, để tính truy thu
  hierarchy: HierarchySnapshot;
  plan: CompPlan;
}

export interface RunOutput {
  lines: EarningLine[];
  warnings: { code: string; message: string; agentCode?: string; policyNo?: string }[];
}

export function runPeriod(input: RunInput): RunOutput;   // hàm thuần, test được bằng Jest
```

Cấu trúc thư mục đề xuất:
```
src/lib/compensation/
  engine/        # types.ts, runPeriod.ts, rules/*.ts (mỗi rule một file), __tests__/
  import/        # parse Excel từ DMS → PremiumEvent[]
  repo/          # đọc/ghi Firestore
src/app/(app)/compensation/
  plans/         # cấu hình chính sách, mô phỏng
  periods/       # chạy kỳ, duyệt, khóa
  statements/    # bảng kê đại lý
```

---

## 6. Màn hình chính

| # | Màn hình | Người dùng | Nội dung chính |
|---|---|---|---|
| 1 | Cấu hình chính sách | Quản lý chính sách | Bảng HH, override, thưởng, truy thu theo phiên bản và ngày hiệu lực; nút "mô phỏng trên kỳ đã khóa" |
| 2 | Import dữ liệu kỳ | Vận hành | Upload Excel, kết quả validate, danh sách lỗi theo dòng |
| 3 | Chạy kỳ và duyệt | Vận hành, Trưởng bộ phận | Tổng hợp theo loại thu nhập, so sánh kỳ trước, cảnh báo trần, nút duyệt / khóa |
| 4 | Bảng kê đại lý | Đại lý, Quản lý | Tổng → loại thu nhập → từng hợp đồng → quy tắc; tải PDF/Excel |
| 5 | Dashboard thu nhập | Đại lý | FYC, APE, số case tháng này, tiến độ tới mốc thưởng, dự kiến thu nhập |
| 6 | Ước tính what-if | Đại lý | Nhập sản phẩm, phí, định kỳ → hoa hồng và thưởng dự kiến |
| 7 | Tra soát | Đại lý, Vận hành | Gửi khiếu nại trên một dòng bảng kê, trạng thái xử lý |
| 8 | Báo cáo tuân thủ | Quản lý | Ngân sách thưởng so với trần Điều 52 theo năm tài chính; hợp đồng chạm trần Điều 51 |

---

## 7. Ví dụ tính thử tháng 10/2026 (số giả định)

Cây: AD Hùng → BM Minh → UM Lan → FC An, FC Bình.

**Sự kiện thu phí trong tháng**
| HĐ | Đại lý | Sản phẩm | Định kỳ | Phí thu | FYC (tỷ lệ năm 1) |
|---|---|---|---|---|---|
| A1 | An | UL | Năm | 24.000.000 | 6.720.000 (28%) |
| A2 | An | UL | Quý | 6.000.000 | 1.680.000 (28%) |
| B1 | Bình | UL | Năm | 40.000.000 | 11.200.000 (28%) |
| B2 | Bình | Tử kỳ | Năm | 5.000.000 | 1.250.000 (25%) |
| B3 | Bình | Bổ trợ SK | Năm | 4.000.000 | 720.000 (18%) |
| L1 | Lan | UL | Năm | 30.000.000 | 8.400.000 (28%) |

**Sự kiện muộn**: HĐ B0 của Bình (UL, phí 10.000.000, phát hành tháng 3/2026) hủy free-look, đã trả FYC 2.800.000 và override UM 560.000, BM 280.000, AD 140.000.

**Kết quả**
| Đại lý | FYC | Override | Thưởng | Truy thu | Thu nhập gộp |
|---|---|---|---|---|---|
| An | 8.400.000 | 0 | 0 (FYC < 10tr) | 0 | 8.400.000 |
| Bình | 13.170.000 | 0 | 1.317.000 (10%, 3 case) | -2.800.000 | 11.687.000 |
| Lan (UM) | 8.400.000 | 4.314.000 (20% × 21.570.000) | 0 (FYC nhóm 21,57tr < 50tr) | -560.000 | 12.154.000 |
| Minh (BM) | 0 | 2.997.000 (10% × 29.970.000) | 0 | -280.000 | 2.717.000 |
| Hùng (AD) | 0 | 1.498.500 (5% × 29.970.000) | 0 | -140.000 | 1.358.500 |

Kiểm tra trần Điều 52 trong tháng: tổng thưởng 1.317.000 so với 30% × 105.000.000 FYP dài hạn = 31.500.000. Đạt. (Trần thật tính lũy kế theo năm tài chính và toàn công ty.)

Thuế TNCN: tính riêng trên dòng cuối bảng kê theo tham số hiện hành, không gộp vào bảng trên.

Mỗi con số trên đều có thể bấm vào để xem: hợp đồng nguồn, sự kiện, quy tắc và tỷ lệ áp dụng. Đó là điểm khác biệt chính so với tính bằng Excel.

---

## 8. Những gì cần thay khi có quy chế thật

1. Toàn bộ bảng trong mục 1 (cấp bậc, tỷ lệ, ngưỡng, lịch truy thu).
2. Định nghĩa APE và K2 theo đúng cách công ty đang tính.
3. Override tính theo cây tại ngày thu phí hay ngày cuối kỳ.
4. Danh sách cột trong file Excel xuất từ DMS để viết bộ import.
5. Quy tắc thuế TNCN đang áp dụng cho đại lý.
