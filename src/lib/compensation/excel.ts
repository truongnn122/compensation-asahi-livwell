import ExcelJS from "exceljs";

import { eventLabel } from "./engine";
import {
  PAYMENT_MODES,
  PREMIUM_EVENT_TYPES,
  type Agent,
  type CompPlan,
  type EarningType,
  type PaidLine,
  type PaymentMode,
  type PremiumEvent,
  type PremiumEventType,
  type RunOutput,
} from "./types";

export const SHEET = {
  agents: "DaiLy",
  events: "ThuPhi",
  paid: "LichSuDaTra",
  guide: "HuongDan",
} as const;

const AGENT_HEADERS = ["Mã ĐL", "Họ tên", "Cấp bậc", "Mã quản lý"];
const EVENT_HEADERS = [
  "Mã sự kiện",
  "Loại sự kiện",
  "Số HĐ",
  "Mã ĐL",
  "Mã sản phẩm",
  "Năm HĐ",
  "Định kỳ đóng",
  "Số tiền",
  "Ngày phát sinh",
  "Ngày phát hành",
];
const PAID_HEADERS = ["Kỳ", "Mã ĐL", "Loại", "Số tiền", "Mã quy tắc", "Số HĐ"];

export const PAYMENT_MODE_LABEL: Record<PaymentMode, string> = {
  ANNUAL: "Năm",
  SEMI: "Nửa năm",
  QUARTERLY: "Quý",
  MONTHLY: "Tháng",
  SINGLE: "Một lần",
};

export const EARNING_TYPE_LABEL: Record<EarningType, string> = {
  FYC: "Hoa hồng năm đầu",
  RYC: "Hoa hồng tái tục",
  OVERRIDE: "Override",
  BONUS: "Thưởng",
  CHARGEBACK: "Truy thu",
};

export type ParsedWorkbook = {
  agents: Agent[];
  events: PremiumEvent[];
  paidHistory: PaidLine[];
  errors: string[];
};

function styleHeader(row: ExcelJS.Row) {
  row.font = { bold: true };
  row.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE6ECF7" },
  };
}

function autoWidth(sheet: ExcelJS.Worksheet, min = 12) {
  sheet.columns.forEach(col => {
    let width = min;
    col.eachCell?.({ includeEmpty: false }, cell => {
      const len = String(cell.value ?? "").length;
      if (len + 2 > width) width = Math.min(len + 2, 48);
    });
    col.width = width;
  });
}

async function toBlob(workbook: ExcelJS.Workbook): Promise<Blob> {
  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

/** Template the user fills from the DMS export. Includes the sample rows passed in. */
export async function buildTemplateWorkbook(
  plan: CompPlan,
  sample: { agents: Agent[]; events: PremiumEvent[]; paidHistory: PaidLine[] }
): Promise<Blob> {
  const wb = new ExcelJS.Workbook();

  const agents = wb.addWorksheet(SHEET.agents);
  styleHeader(agents.addRow(AGENT_HEADERS));
  for (const a of sample.agents) {
    agents.addRow([a.code, a.name, a.rank, a.managerCode]);
  }
  autoWidth(agents);

  const events = wb.addWorksheet(SHEET.events);
  styleHeader(events.addRow(EVENT_HEADERS));
  for (const e of sample.events) {
    events.addRow([
      e.eventId,
      e.type,
      e.policyNo,
      e.agentCode,
      e.productCode,
      e.policyYear,
      e.paymentMode,
      e.amount,
      e.occurredAt,
      e.issueDate,
    ]);
  }
  autoWidth(events);

  const paid = wb.addWorksheet(SHEET.paid);
  styleHeader(paid.addRow(PAID_HEADERS));
  for (const p of sample.paidHistory) {
    paid.addRow([
      p.periodId,
      p.agentCode,
      p.type,
      p.amount,
      p.ruleId,
      p.sourcePolicyNo,
    ]);
  }
  autoWidth(paid);

  const guide = wb.addWorksheet(SHEET.guide);
  styleHeader(guide.addRow(["Mục", "Giá trị hợp lệ", "Ghi chú"]));
  guide.addRow([
    "Cấp bậc",
    plan.ranks.map(r => `${r.code} = ${r.name}`).join("; "),
    "Cột Cấp bậc trong sheet DaiLy",
  ]);
  guide.addRow([
    "Mã sản phẩm",
    plan.products.map(p => `${p.code} = ${p.name}`).join("; "),
    "Cột Mã sản phẩm trong sheet ThuPhi",
  ]);
  guide.addRow([
    "Loại sự kiện",
    PREMIUM_EVENT_TYPES.map(t => `${t} = ${eventLabel(t)}`).join("; "),
    "PREMIUM_COLLECTED sinh hoa hồng; các loại còn lại sinh truy thu",
  ]);
  guide.addRow([
    "Định kỳ đóng",
    PAYMENT_MODES.map(m => `${m} = ${PAYMENT_MODE_LABEL[m]}`).join("; "),
    "Dùng để quy đổi APE",
  ]);
  guide.addRow([
    "Ngày",
    "YYYY-MM-DD hoặc ô định dạng ngày của Excel",
    "Ngày phát sinh là ngày thu phí thực tế",
  ]);
  guide.addRow(["Số tiền", "VND, số dương", "Hoa hồng tính trên phí thực thu"]);
  guide.addRow([
    "LichSuDaTra",
    "Các khoản đã trả ở kỳ trước cho HĐ bị hủy / mất hiệu lực",
    "Dùng để tính truy thu. Bỏ trống nếu không có",
  ]);
  autoWidth(guide, 16);

  return toBlob(wb);
}

function cellText(cell: ExcelJS.Cell): string {
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return toIsoDate(v);
  if (typeof v === "object") {
    if ("result" in v && v.result !== undefined) return String(v.result);
    if ("richText" in v) return v.richText.map(r => r.text).join("");
    if ("text" in v) return String(v.text);
  }
  return String(v).trim();
}

function toIsoDate(d: Date): string {
  const pad = (n: number) => n.toString().padStart(2, "0");
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
}

function parseDate(raw: string): string | null {
  if (!raw) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return raw.slice(0, 10);
  const dmy = raw.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/);
  if (dmy) {
    const [, d, m, y] = dmy;
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : toIsoDate(parsed);
}

function parseNumber(raw: string): number | null {
  if (!raw) return null;
  const n = Number(raw.replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n : null;
}

function rowValues(row: ExcelJS.Row, count: number): string[] {
  const out: string[] = [];
  for (let i = 1; i <= count; i++) out.push(cellText(row.getCell(i)));
  return out;
}

function isBlank(values: string[]) {
  return values.every(v => v === "");
}

export async function parseWorkbook(
  data: ArrayBuffer,
  plan: CompPlan
): Promise<ParsedWorkbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(data);
  const errors: string[] = [];
  const agents: Agent[] = [];
  const events: PremiumEvent[] = [];
  const paidHistory: PaidLine[] = [];
  const rankCodes = new Set(plan.ranks.map(r => r.code));
  const productCodes = new Set(plan.products.map(p => p.code));

  const agentSheet = wb.getWorksheet(SHEET.agents);
  if (!agentSheet) errors.push(`Thiếu sheet "${SHEET.agents}".`);
  else {
    agentSheet.eachRow((row, idx) => {
      if (idx === 1) return;
      const [code, name, rank, managerCode] = rowValues(row, 4);
      if (isBlank([code, name, rank, managerCode])) return;
      if (!code) errors.push(`${SHEET.agents} dòng ${idx}: thiếu Mã ĐL.`);
      if (!rankCodes.has(rank))
        errors.push(
          `${SHEET.agents} dòng ${idx}: cấp bậc "${rank}" không có trong chính sách.`
        );
      agents.push({ code, name: name || code, rank, managerCode });
    });
  }
  const agentCodes = new Set(agents.map(a => a.code));
  for (const a of agents) {
    if (a.managerCode && !agentCodes.has(a.managerCode))
      errors.push(
        `${SHEET.agents}: đại lý ${a.code} có mã quản lý "${a.managerCode}" không tồn tại.`
      );
  }

  const eventSheet = wb.getWorksheet(SHEET.events);
  if (!eventSheet) errors.push(`Thiếu sheet "${SHEET.events}".`);
  else {
    const seen = new Set<string>();
    eventSheet.eachRow((row, idx) => {
      if (idx === 1) return;
      const values = rowValues(row, 10);
      if (isBlank(values)) return;
      const [
        eventIdRaw,
        typeRaw,
        policyNo,
        agentCode,
        productCode,
        yearRaw,
        modeRaw,
        amountRaw,
        occurredRaw,
        issueRaw,
      ] = values;
      const where = `${SHEET.events} dòng ${idx}`;
      const type = (typeRaw || "PREMIUM_COLLECTED") as PremiumEventType;
      if (!PREMIUM_EVENT_TYPES.includes(type))
        errors.push(`${where}: loại sự kiện "${typeRaw}" không hợp lệ.`);
      if (!policyNo) errors.push(`${where}: thiếu Số HĐ.`);
      if (!agentCodes.has(agentCode))
        errors.push(
          `${where}: Mã ĐL "${agentCode}" không có trong sheet DaiLy.`
        );
      if (!productCodes.has(productCode))
        errors.push(
          `${where}: mã sản phẩm "${productCode}" chưa có trong chính sách.`
        );
      const policyYear = parseNumber(yearRaw) ?? 1;
      const paymentMode = (modeRaw || "ANNUAL") as PaymentMode;
      if (!PAYMENT_MODES.includes(paymentMode))
        errors.push(`${where}: định kỳ đóng "${modeRaw}" không hợp lệ.`);
      const amount = parseNumber(amountRaw);
      if (amount === null || amount <= 0)
        errors.push(`${where}: số tiền không hợp lệ.`);
      const occurredAt = parseDate(occurredRaw);
      if (!occurredAt) errors.push(`${where}: ngày phát sinh không hợp lệ.`);
      const issueDate = parseDate(issueRaw) ?? occurredAt;
      const eventId = eventIdRaw || `R${idx}`;
      if (seen.has(eventId))
        errors.push(`${where}: mã sự kiện "${eventId}" bị trùng.`);
      seen.add(eventId);
      events.push({
        eventId,
        type,
        policyNo,
        agentCode,
        productCode,
        policyYear,
        amount: amount ?? 0,
        occurredAt: occurredAt ?? "",
        issueDate: issueDate ?? "",
        paymentMode,
      });
    });
  }

  const paidSheet = wb.getWorksheet(SHEET.paid);
  if (paidSheet) {
    paidSheet.eachRow((row, idx) => {
      if (idx === 1) return;
      const values = rowValues(row, 6);
      if (isBlank(values)) return;
      const [periodId, agentCode, typeRaw, amountRaw, ruleId, policyNo] =
        values;
      const amount = parseNumber(amountRaw);
      if (amount === null)
        errors.push(`${SHEET.paid} dòng ${idx}: số tiền không hợp lệ.`);
      paidHistory.push({
        periodId,
        agentCode,
        type: (typeRaw || "FYC") as EarningType,
        amount: amount ?? 0,
        ruleId: ruleId || "",
        sourcePolicyNo: policyNo,
      });
    });
  }

  return { agents, events, paidHistory, errors };
}

/** Results workbook: summary, every earning line, warnings. */
export async function buildResultWorkbook(
  output: RunOutput,
  plan: CompPlan
): Promise<Blob> {
  const wb = new ExcelJS.Workbook();
  const rankName = new Map(plan.ranks.map(r => [r.code, r.name]));
  const money = "#,##0";

  const summary = wb.addWorksheet("TongHop");
  styleHeader(
    summary.addRow([
      "Kỳ",
      "Mã ĐL",
      "Họ tên",
      "Cấp bậc",
      "Số HĐ",
      "FYP",
      "APE",
      "Hoa hồng năm đầu",
      "Hoa hồng tái tục",
      "Override",
      "Thưởng",
      "Truy thu",
      "Thu nhập gộp",
    ])
  );
  for (const s of output.summaries) {
    summary.addRow([
      output.periodId,
      s.agent.code,
      s.agent.name,
      rankName.get(s.agent.rank) ?? s.agent.rank,
      s.cases,
      s.fyp,
      s.ape,
      s.fyc,
      s.ryc,
      s.override,
      s.bonus,
      s.chargeback,
      s.gross,
    ]);
  }
  const t = output.totals;
  const totalRow = summary.addRow([
    output.periodId,
    "",
    "Tổng",
    "",
    t.cases,
    t.fyp,
    t.ape,
    t.fyc,
    t.ryc,
    t.override,
    t.bonus,
    t.chargeback,
    t.gross,
  ]);
  totalRow.font = { bold: true };
  for (let c = 6; c <= 13; c++) summary.getColumn(c).numFmt = money;
  autoWidth(summary);

  const detail = wb.addWorksheet("ChiTiet");
  styleHeader(
    detail.addRow([
      "Kỳ",
      "Mã ĐL",
      "Họ tên",
      "Loại",
      "Số HĐ",
      "Mã quy tắc",
      "Sự kiện nguồn",
      "Diễn giải",
      "Số tiền",
    ])
  );
  const nameOf = new Map(
    output.summaries.map(s => [s.agent.code, s.agent.name])
  );
  for (const l of output.lines) {
    detail.addRow([
      l.periodId,
      l.agentCode,
      nameOf.get(l.agentCode) ?? "",
      EARNING_TYPE_LABEL[l.type],
      l.sourcePolicyNo,
      l.ruleId,
      l.sourceEventId,
      l.formula,
      l.amount,
    ]);
  }
  detail.getColumn(9).numFmt = money;
  autoWidth(detail);

  const warn = wb.addWorksheet("CanhBao");
  styleHeader(warn.addRow(["Mã", "Mức", "Nội dung", "Mã ĐL", "Số HĐ"]));
  for (const w of output.warnings) {
    warn.addRow([
      w.code,
      w.level,
      w.message,
      w.agentCode ?? "",
      w.policyNo ?? "",
    ]);
  }
  warn.addRow([
    "R-CAP-52",
    "info",
    `Trần thưởng Điều 52: ${Math.round(output.bonusCap.cap).toLocaleString("vi-VN")} · đã dùng ${Math.round(output.bonusCap.used).toLocaleString("vi-VN")}`,
    "",
    "",
  ]);
  autoWidth(warn, 14);

  return toBlob(wb);
}
