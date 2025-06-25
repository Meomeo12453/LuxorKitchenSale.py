import pandas as pd

# Đọc dữ liệu gốc
file_in = "monthly_report_20250501.xlsx"  # thay bằng file của bạn nếu khác tên
sheet = "Table 4"
df = pd.read_excel(file_in, sheet_name=sheet)

# Danh sách các mục phí muốn phân tích
fee_items = [
    'Phí cố định',
    'Phí thanh toán',
    'Phí hoa hồng Tiếp thị liên kết',
    'Phí dịch vụ PiShip'
]

# Lọc và lấy giá trị từng loại phí (dấu âm là chi phí)
df_fees = df[df['Tổng kết thanh toán đã chuyển'].astype(str).str.contains("Phí", na=False)]
df_fees = df_fees.set_index('Tổng kết thanh toán đã chuyển')
fee_values = []
for item in fee_items:
    try:
        val = float(df_fees.loc[item, 'Số tiền (VND)'])
        fee_values.append(abs(val))
    except Exception:
        fee_values.append(0)

# Lấy tổng doanh thu sản phẩm
revenue = float(df[df['Tổng kết thanh toán đã chuyển'] == "Tổng tiền sản phẩm"]['Số tiền (VND)'].values[0])
# Tổng phí
total_fees = sum(fee_values)

# Tạo DataFrame xuất Excel
result = pd.DataFrame({
    'Tên chi phí': fee_items + ['Tổng phí giao dịch', 'Tổng doanh thu sản phẩm', 'Tỷ lệ phí/doanh thu (%)'],
    'Giá trị (VND)': fee_values + [total_fees, revenue, ''],
    'Tỷ lệ so với doanh thu (%)': [f"{v/revenue*100:.2f}" for v in fee_values] + [f"{total_fees/revenue*100:.2f}", '', f"{total_fees/revenue*100:.2f}"]
})

# Xuất file Excel
file_out = "bao_cao_chi_phi.xlsx"
result.to_excel(file_out, index=False)
print("Đã xuất file:", file_out)
