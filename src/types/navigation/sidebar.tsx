import {
  LayoutDashboard,
  type LucideIcon,
  UserCheck,
  Users,
} from "lucide-react";

import { canAccessUsers, type Role } from "@/lib/permissions";

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

  if (canAccessUsers(role)) {
    items.push({ title: "Người dùng", url: "/users", icon: Users });
  }

  return [{ id: 1, items }];
}
