"use client";
import * as React from "react";

import {
  ChartBar,
  Forklift,
  Gauge,
  GraduationCap,
  LayoutDashboard,
  Search,
  ShoppingBag,
} from "lucide-react";

import { Button } from "@/components/ui/button";
import {
  CommandDialog,
  CommandEmpty,
  CommandGroup,
  CommandInput,
  CommandItem,
  CommandList,
  CommandSeparator,
} from "@/components/ui/command";

const searchItems = [
  { group: "Bảng điều khiển", icon: LayoutDashboard, label: "Mặc định" },
  { group: "Bảng điều khiển", icon: ChartBar, label: "CRM", disabled: true },
  {
    group: "Bảng điều khiển",
    icon: Gauge,
    label: "Phân tích",
    disabled: true,
  },
  {
    group: "Bảng điều khiển",
    icon: ShoppingBag,
    label: "Thương mại điện tử",
    disabled: true,
  },
  {
    group: "Bảng điều khiển",
    icon: GraduationCap,
    label: "Học viện",
    disabled: true,
  },
  {
    group: "Bảng điều khiển",
    icon: Forklift,
    label: "Hậu cần",
    disabled: true,
  },
  { group: "Xác thực", label: "Đăng nhập v1" },
  { group: "Xác thực", label: "Đăng nhập v2" },
  { group: "Xác thực", label: "Đăng ký v1" },
  { group: "Xác thực", label: "Đăng ký v2" },
];

export function SearchDialog() {
  const [open, setOpen] = React.useState(false);
  React.useEffect(() => {
    const down = (e: KeyboardEvent) => {
      if (e.key === "j" && (e.metaKey || e.ctrlKey)) {
        e.preventDefault();
        setOpen(open => !open);
      }
    };
    document.addEventListener("keydown", down);
    return () => document.removeEventListener("keydown", down);
  }, []);

  return (
    <>
      <Button
        variant="link"
        className="text-muted-foreground !px-0 font-normal hover:no-underline"
        onClick={() => setOpen(true)}
      >
        <Search className="size-4" />
        Tìm kiếm
        <kbd className="bg-muted inline-flex h-5 items-center gap-1 rounded border px-1.5 text-[10px] font-medium select-none">
          <span className="text-xs">⌘</span>J
        </kbd>
      </Button>
      <CommandDialog open={open} onOpenChange={setOpen}>
        <CommandInput placeholder="Tìm kiếm bảng điều khiển, người dùng, và hơn thế nữa…" />
        <CommandList>
          <CommandEmpty>Không tìm thấy kết quả.</CommandEmpty>
          {[...new Set(searchItems.map(item => item.group))].map((group, i) => (
            <React.Fragment key={group}>
              {i !== 0 && <CommandSeparator />}
              <CommandGroup heading={group} key={group}>
                {searchItems
                  .filter(item => item.group === group)
                  .map(item => (
                    <CommandItem
                      className="!py-1.5"
                      key={item.label}
                      onSelect={() => setOpen(false)}
                    >
                      {item.icon && <item.icon />}
                      <span>{item.label}</span>
                    </CommandItem>
                  ))}
              </CommandGroup>
            </React.Fragment>
          ))}
        </CommandList>
      </CommandDialog>
    </>
  );
}
