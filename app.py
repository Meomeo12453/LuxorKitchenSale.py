# 0. Upload file, tự động chuẩn hóa tên cột không dấu và in ra đường dẫn
from google.colab import files
import pandas as pd
import matplotlib.pyplot as plt
from IPython.display import display, HTML
import unicodedata
import re

def chuan_hoa_ten_cot(s):
    # Xóa dấu, xóa ký tự đặc biệt, chuyển thường, thay khoảng trắng thành "_"
    s = unicodedata.normalize('NFD', s)
    s = s.encode('ascii', 'ignore').decode('utf-8')
    s = re.sub(r'[^\w\s]', '', s)
    s = s.strip().lower().replace(' ', '_')
    return s

uploaded = files.upload()
if not uploaded:
    raise FileNotFoundError("Bạn chưa upload file nào.")
file_name = list(uploaded.keys())[0]
excel_file_path = f'/content/{file_name}'
print(f"✅ File Excel đã upload: {excel_file_path}")

# Tùy chọn giao diện
hide_code = True
chart_to_show = "All"  # Hoặc "Bar", "Box", "Pie"

# 1. CSS full-width & ẩn code nếu cần
css = f"""
<style>
  .container {{ width:100% !important; max-width:100% !important; }}
  {' .input { display: none; }' if hide_code else ''}
</style>
"""
display(HTML(css))

# 2. Đọc dữ liệu Excel và chuẩn hóa tên cột
try:
    df = pd.read_excel(excel_file_path)
    df.columns = [chuan_hoa_ten_cot(col) for col in df.columns]
    display(HTML(f"<p style='color:green;'>✅ Đã đọc file: <code>{excel_file_path}</code></p>"))
except Exception as e:
    raise RuntimeError(f"❌ Lỗi khi đọc Excel: {e}")

# Hiển thị tên cột mẫu để kiểm tra
print("Các cột sau khi chuẩn hóa:", list(df.columns))

# 3. Định nghĩa cấu trúc hoa hồng & tính toán override_sales
network = {
    'catalyst':     {'comm_rate': 0.35, 'override_rate': 0.00},
    'visionary':    {'comm_rate': 0.40, 'override_rate': 0.05},
    'trailblazer':  {'comm_rate': 0.40, 'override_rate': 0.05},
}

# Đảm bảo các cột cần thiết luôn có
required_cols = ['ma_khach_hang', 'nhom_khach_hang', 'tong_ban_tru_tra_hang', 'ghi_chu']
for col in required_cols:
    if col not in df.columns:
        raise ValueError(f"Thiếu cột: {col}. Vui lòng kiểm tra file Excel.")

def calculate_override(df_in):
    df2 = df_in.copy()
    df2['override_sales'] = 0
    for role in network:
        staff = df2[df2['nhom_khach_hang'].str.lower() == role]
        for _, row in staff.iterrows():
            mask = df2['ghi_chu'] == row['ma_khach_hang']
            subtotal = df2.loc[mask, 'tong_ban_tru_tra_hang'].sum()
            df2.loc[df2['ma_khach_hang'] == row['ma_khach_hang'], 'override_sales'] = subtotal
    return df2

df = calculate_override(df)
df['comm_rate']     = df['nhom_khach_hang'].str.lower().map(lambda r: network[r]['comm_rate'] if r in network else 0)
df['override_rate'] = df['nhom_khach_hang'].str.lower().map(lambda r: network[r]['override_rate'] if r in network else 0)
df['override_comm'] = df['override_sales'] * df['override_rate']

# Hiển thị 5 dòng đầu để kiểm tra
display(df.head())

# 4. Vẽ biểu đồ
def plot_bar():
    plt.figure(figsize=(8,4))
    plt.bar(df['ma_khach_hang'], df['tong_ban_tru_tra_hang'])
    plt.title('Doanh số theo khách hàng')
    plt.xlabel('Mã khách hàng')
    plt.ylabel('Tổng bán (trừ trả hàng)')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

def plot_box():
    plt.figure(figsize=(8,4))
    data = [df[df['nhom_khach_hang'].str.lower()==r]['tong_ban_tru_tra_hang'] for r in network]
    plt.boxplot(data, labels=[r.capitalize() for r in network])
    plt.title('Phân phối doanh số theo nhóm khách hàng')
    plt.tight_layout()
    plt.show()

def plot_pie():
    plt.figure(figsize=(6,6))
    s = df.groupby(df['nhom_khach_hang'].str.capitalize())['tong_ban_tru_tra_hang'].sum()
    plt.pie(s, labels=s.index, autopct='%1.1f%%')
    plt.title('Tỷ trọng doanh số theo nhóm khách hàng')
    plt.axis('equal')
    plt.show()

if chart_to_show in ("All","Bar"): plot_bar()
if chart_to_show in ("All","Box"): plot_box()
if chart_to_show in ("All","Pie"): plot_pie()

# 5. Xuất báo cáo ra Excel
output_file = '/content/sales_report_tong_hop.xlsx'
df.to_excel(output_file, index=False)
print(f"Đã lưu file kết quả: {output_file}")

# 6. Download file ngay trên Colab
files.download(output_file)
