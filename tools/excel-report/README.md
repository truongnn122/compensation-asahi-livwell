# Tool báo cáo thu nhập Đại lý (Python → Excel)

## Cài đặt
```bash
pip install pandas openpyxl
```

## Sử dụng
1. Tạo file mẫu: `python tao_file_mau.py` → `dau_vao_mau.xlsx`
2. Dán dữ liệu thật vào sheet **DuLieuDaiLy**, sửa tham số ở sheet **ChinhSach** (ô vàng).
3. Chạy: `python bao_cao_thu_nhap.py dau_vao_mau.xlsx --ky T10-2026`
4. Mở `bao_cao_thu_nhap_T10-2026.xlsx`.

## Logic tính (mẫu – cần thay bằng chính sách thật)
| Khoản | Công thức |
|---|---|
| FYC | FYP × tỷ lệ hoa hồng |
| Điều kiện thưởng | Số HĐ ≥ tối thiểu **và** K2 ≥ tối thiểu |
| Thưởng hoạt động | FYC × tỷ lệ theo bậc FYC (nếu đạt ĐK) |
| Tổng trước thuế | FYC + Thưởng + Thu nhập khác |
| Thực nhận | Tổng trước thuế × (1 − thuế TNCN) |

Mọi cột tính trong báo cáo là công thức Excel → sửa sheet ChinhSach trong file kết quả là số tự tính lại.
