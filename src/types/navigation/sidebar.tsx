import {
  FileText,
  LayoutDashboard,
  type LucideIcon,
  Settings,
} from "lucide-react";

export interface NavSubItem {
  title: string;
  url: string;
  icon?: LucideIcon;
  comingSoon?: boolean;
  newTab?: boolean;
  isNew?: boolean;
}

export interface NavMainItem {
  title: string;
  url: string;
  icon?: LucideIcon;
  subItems?: NavSubItem[];
  comingSoon?: boolean;
  newTab?: boolean;
  isNew?: boolean;
}

export interface NavGroup {
  id: number;
  label?: string;
  items: NavMainItem[];
}

export const sidebarItems: NavGroup[] = [
  {
    id: 1,
    items: [
      { title: "Bảng điều khiển", url: "/dashboard", icon: LayoutDashboard },
      { title: "Tài liệu", url: "/documents", icon: FileText },
      { title: "Cài đặt", url: "/settings", icon: Settings },
    ],
  },
];
