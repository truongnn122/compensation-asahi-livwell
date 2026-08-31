import {
  FileText,
  LayoutDashboard,
  type LucideIcon,
  Settings,
  UserCheck,
  Users,
} from "lucide-react";

import {
  canAccessDocuments,
  canAccessSettings,
  canAccessUsers,
  type Role,
} from "@/lib/permissions";

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

export function getSidebarItems(role: Role): NavGroup[] {
  const items: NavMainItem[] = [
    { title: "Bảng điều khiển", url: "/dashboard", icon: LayoutDashboard },
    { title: "Ứng viên tuyển dụng", url: "/recruitments", icon: UserCheck },
  ];

  if (canAccessDocuments(role)) {
    items.push({ title: "Tài liệu", url: "/documents", icon: FileText });
  }
  if (canAccessUsers(role)) {
    items.push({ title: "Người dùng", url: "/users", icon: Users });
  }
  if (canAccessSettings(role)) {
    items.push({ title: "Cài đặt", url: "/settings", icon: Settings });
  }

  return [{ id: 1, items }];
}
