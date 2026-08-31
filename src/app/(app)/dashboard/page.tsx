import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Bảng điều khiển",
};

export default function DashboardPage() {
  return <h1 className="text-2xl font-semibold">Bảng điều khiển</h1>;
}
