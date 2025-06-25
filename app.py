import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import tkinter as tk
from tkinter import filedialog

# === BƯỚC 1: Chọn file qua hộp thoại (không cần cùng thư mục) ===
root = tk.Tk()
root.withdraw()
file_in = filedialog.askopenfilename(title="Chọn file Excel", filetypes=[("Excel files", "*.xlsx")])
if not file_in:
    raise Exception("Bạn chưa chọn file!")

sheet = "Table 4"
df = pd.read_excel(file_in, sheet_name=sheet)

# === BƯỚC 2: Xử lý phí & doanh thu ===
try:
    revenue = float(df[df['Tổng kết thanh toán đã chuyển'] == "Tổng tiền sản phẩm"]['Số tiền (VND)'].values[0])
except:
    raise Exception("Không tìm thấy dòng 'Tổng tiền sản phẩm' trong sheet Table 4!")

fees_df = df[df['Tổng kết thanh toán đã chuyển'].astype(str).str.contains('Phí', na=False)].copy()
fees_df['Giá trị (VND)'] = fees_df['Số tiền (VND)'].astype(float).abs()
fees_df['Tỷ lệ so với doanh thu (%)'] = fees_df['Giá trị (VND)'] / revenue * 100

fee_names = fees_df['Tổng kết thanh toán đã chuyển'].tolist()
fee_values = fees_df['Giá trị (VND)'].tolist()
fee_percent = fees_df['Tỷ lệ so với doanh thu (%)'].tolist()
total_fees = sum(fee_values)

# === BƯỚC 3: Bar chart phần trăm từng loại phí ===
plt.figure(figsize=(10,5))
plt.bar(fee_names, fee_percent, color='orange')
plt.ylabel("Phần trăm (%)")
plt.title("Tỷ lệ phần trăm từng loại chi phí so với tổng doanh thu sản phẩm")
plt.xticks(rotation=25, ha='right')
plt.tight_layout()
plt.show()

# === BƯỚC 4: Pie chart tỷ trọng phí ===
plt.figure(figsize=(8,8))
plt.pie(fee_values, labels=fee_names, autopct='%1.1f%%', startangle=90)
plt.title("Tỷ trọng từng loại phí giao dịch")
plt.axis('equal')
plt.tight_layout()
plt.show()

# === BƯỚC 5: Waterfall chart (Plotly) ===
labels = ['Tổng tiền sản phẩm'] + fee_names + ['Sau khi trừ phí']
values = [revenue] + [-v for v in fee_values] + [revenue - total_fees]
measure = ["absolute"] + ["relative"]*len(fee_names) + ["total"]
fig = go.Figure(go.Waterfall(
    name="Dòng tiền",
    orientation="v",
    measure=measure,
    x=labels,
    y=values,
    text=[f"{abs(v):,.0f}₫" for v in values],
    connector={"line": {"color": "orange"}}
))
fig.update_layout(title="Biểu đồ thác nước: Dòng tiền trên tổng doanh thu sản phẩm", height=500)
fig.show()

# === BƯỚC 6: Xuất file Excel tổng hợp ===
result = fees_df[['Tổng kết thanh toán đã chuyển', 'Giá trị (VND)', 'Tỷ lệ so với doanh thu (%)']].copy()
result = result.rename(columns={'Tổng kết thanh toán đã chuyển': 'Tên chi phí'})
result = result.append({
    'Tên chi phí': 'Tổng phí giao dịch',
    'Giá trị (VND)': total_fees,
    'Tỷ lệ so với doanh thu (%)': total_fees / revenue * 100
}, ignore_index=True)
result = result.append({
    'Tên chi phí': 'Tổng doanh thu sản phẩm',
    'Giá trị (VND)': revenue,
    'Tỷ lệ so với doanh thu (%)': ''
}, ignore_index=True)
result['Tỷ lệ so với doanh thu (%)'] = result['Tỷ lệ so với doanh thu (%)'].apply(lambda x: f"{x:.2f}%" if x != '' else '')

output_excel = "bao_cao_chi_phi.xlsx"
result.to_excel(output_excel, index=False)
print(f"✅ Đã xuất file tổng hợp: {output_excel}")
