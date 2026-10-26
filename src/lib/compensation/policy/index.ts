/**
 * MVI agency compensation policy engine, ported from the Income Tool
 * (memo set ALL/AD/2608). Every function is pure and synchronous; money is
 * in triệu VND and percentages are 0–100 numbers. `k2Percent: null` means
 * "chưa có K2 / N/A".
 */
export * from "./commission";
export * from "./fixedAllowance";
export * from "./monthlyBonus";
export * from "./quarterlyBonus";
export * from "./mdrtBonus";
export * from "./healthInsurance";
export * from "./managerAllowance";
export * from "./managerBonus";
export * from "./indirectSupport";
export * from "./departmentBonus";
export * from "./mdrtBuyout";
