import streamlit as st
from PIL import Image
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.express as px
import colorsys
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font

# ===== Cấu hình giao diện =====
st.set_page_config(page_title="Sales Dashboard MiniApp", layout="wide")
st.markdown("""
    <style>
    .block-container {padding-top:1.2rem;}
    .stApp {background: #F7F8FA;}
    img { border-radius: 0 !important; }
    </style>
    """, unsafe_allow_html=True)

# ==== Hiển thị logo căn giữa ====
LOGO_PATHS = [
    "logo-daba.png",
    "ef5ac011-857d-4b32-bd70-ef9ac3817106.png",
    "30313609-d84b-45c1-958e-7d50bf11b60c.png",
    "002f43d6-a413-41d0-b88a-cde6a1a1a98c.png"
]
logo = None
for path in LOGO_PATHS:
    if os.path.exists(path):
        logo = Image.open(path)
        break

if logo is not None:
    desired_height = 36  # pixel
    w, h = logo.size
    new_width = int((w / h) * desired_height)
    logo_resized = logo.resize((new_width, desired_height))
    st.markdown("<div style='display:flex;justify-content:center;margin-bottom:12px;'>", unsafe_allow_html=True)
    st.image(logo_resized)
    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.warning("Không tìm thấy file logo. Đảm bảo file logo đã upload đúng thư mục app!")

# ===== HOTLINE & ĐỊA CHỈ =====
st.markdown(
    "<div style='text-align:center;font-size:16px;color:#1570af;font-weight:600;'>Hotline: 0909.625.808</div>",
    unsafe_allow_html=True)
st.markdown(
    "<div style='text-align:center;font-size:14px;color:#555;'>Địa chỉ: Lầu 9, Pearl Plaza, 561A Điện Biên Phủ, P.25, Q. Bình Thạnh, TP.HCM</div>",
    unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

# ===== TIÊU ĐỀ =====
st.title("Sales Dashboard MiniApp")
st.markdown(
    "<small style='color:gray;'>Dashboard phân tích & quản trị đại lý cho DABA Sài Gòn. Tải file Excel, lọc – tra cứu – trực quan – tải báo cáo màu nhóm.</small>",
    unsafe_allow_html=True)

# ======= CÁC CONTROL PHÂN TÍCH TRÊN MAIN PAGE =======
st.markdown("## 🔎 Tùy chọn phân tích")
col1, col2 = st.columns(2)
with col1:
    chart_type = st.radio("Chọn biểu đồ:", ["Cột chồng", "Sunburst", "Pareto", "Pie"], horizontal=True)
with col2:
    filter_nganh = st.multiselect("Lọc theo nhóm khách hàng:", ["Catalyst", "Visionary", "Trailblazer"], default=[])

st.markdown("### 1. Tải lên file Excel (.xlsx)")
uploaded_file = st.file_uploader("", type="xlsx", help="Chỉ nhận Excel, <200MB.")
if not uploaded_file:
    st.info("💡 Hãy upload file Excel mẫu để bắt đầu sử dụng Dashboard.")
    with st.expander("📋 Xem hướng dẫn & file mẫu", expanded=False):
        st.markdown(
            "- Nhấn **Browse files** hoặc kéo thả file.\n"
            "- File cần các cột: **Mã khách hàng, Tên khách hàng, Nhóm khách hàng, Tổng bán trừ trả hàng, Ghi chú**.\n"
            "- Nếu lỗi, kiểm tra lại tiêu đề cột trong file Excel."
        )
    st.stop()

# ===== XỬ LÝ DỮ LIỆU =====
df = pd.read_excel(uploaded_file)
df['Mã khách hàng'] = df['Mã khách hàng'].astype(str)

# ---- Xác định parent-child logic mới (cấp trên dựa vào Ghi chú) ----
parent_ids = []
prev_ma_khach_hang = set()
for idx, row in df.iterrows():
    parent_id = None
    if pd.notnull(row['Ghi chú']) and row['Ghi chú'] in prev_ma_khach_hang:
        parent_id = row['Ghi chú']
    parent_ids.append(parent_id)
    prev_ma_khach_hang.add(row['Mã khách hàng'])
df['parent_id'] = parent_ids

# ---- Số thuộc cấp trực tiếp (F1) ----
f1_counts = []
for ma_kh in df['Mã khách hàng']:
    f1_count = (df['parent_id'] == ma_kh).sum()
    f1_counts.append(f1_count)
df['Số thuộc cấp F1'] = f1_counts

# ---- Tổng doanh số F1 (tổng doanh số của các cấp dưới trực tiếp) ----
ds_f1 = []
for ma_kh in df['Mã khách hàng']:
    subtotal = df.loc[df['parent_id'] == ma_kh, 'Tổng bán trừ trả hàng'].sum()
    ds_f1.append(subtotal)
df['Doanh số F1'] = ds_f1

# ---- Cấu trúc hoa hồng (tùy chỉnh theo từng nhóm) ----
network = {
    'Catalyst':     {'comm_rate': 0.35, 'override_rate': 0.00},
    'Visionary':    {'comm_rate': 0.40, 'override_rate': 0.05},
    'Trailblazer':  {'comm_rate': 0.40, 'override_rate': 0.05},
}
df['comm_rate']     = df['Nhóm khách hàng'].map(lambda r: network.get(r, {}).get('comm_rate', 0))
df['override_rate'] = df['Nhóm khách hàng'].map(lambda r: network.get(r, {}).get('override_rate', 0))
df['override_comm'] = df['Doanh số F1'] * df['override_rate']

# ---- Filter nhóm ----
if filter_nganh:
    df = df[df['Nhóm khách hàng'].isin(filter_nganh)]

# ===== BẢNG DỮ LIỆU & GIẢI THÍCH =====
with st.expander("📋 Giải thích các trường dữ liệu", expanded=False):
    st.markdown("""
    **Các trường dữ liệu chính:**  
    - `parent_id`: Mã khách hàng cấp trên trực tiếp (nếu có).
    - `Số thuộc cấp F1`: Số thành viên trực tiếp dưới nhánh này.
    - `Doanh số F1`: Tổng doanh số của các cấp dưới trực tiếp.
    - `override_comm`: Hoa hồng từ hệ thống cấp dưới F1 (áp dụng tỷ lệ từng nhóm).
    """)

st.subheader("2. Bảng dữ liệu đại lý đã xử lý")
st.dataframe(df, use_container_width=True, hide_index=True)

# ===== BIỂU ĐỒ PHÂN TÍCH =====
st.subheader("3. Biểu đồ phân tích dữ liệu")

if chart_type == "Cột chồng":
    fig, ax = plt.subplots(figsize=(12,5))
    ind = np.arange(len(df))
    ax.bar(ind, df['Tổng bán trừ trả hàng'], width=0.5, label='Tổng bán cá nhân')
    ax.bar(ind, df['override_comm'], width=0.5, bottom=df['Tổng bán trừ trả hàng'], label='Hoa hồng hệ thống')
    ax.set_ylabel('Số tiền (VND)')
    ax.set_title('Tổng bán & Hoa hồng hệ thống từng cá nhân')
    ax.set_xticks(ind)
    ax.set_xticklabels(df['Tên khách hàng'], rotation=60, ha='right')
    ax.legend()
    st.pyplot(fig)

elif chart_type == "Sunburst":
    try:
        fig2 = px.sunburst(
            df,
            path=['Nhóm khách hàng', 'Tên khách hàng'],
            values='Tổng bán trừ trả hàng',
            title="Sơ đồ hệ thống cấp bậc & doanh số"
        )
        st.plotly_chart(fig2, use_container_width=True)
    except Exception as e:
        st.error(f"Lỗi khi vẽ Sunburst chart: {e}")

elif chart_type == "Pareto":
    try:
        df_sorted = df.sort_values('Tổng bán trừ trả hàng', ascending=False)
        cum_sum = df_sorted['Tổng bán trừ trả hàng'].cumsum()
        cum_perc = 100 * cum_sum / df_sorted['Tổng bán trừ trả hàng'].sum()
        fig3, ax1 = plt.subplots(figsize=(10,5))
        ax1.bar(np.arange(len(df_sorted)), df_sorted['Tổng bán trừ trả hàng'], label="Doanh số")
        ax1.set_ylabel('Doanh số')
        ax1.set_xticks(range(len(df_sorted)))
        ax1.set_xticklabels(df_sorted['Tên khách hàng'], rotation=60, ha='right')
        ax2 = ax1.twinx()
        ax2.plot(np.arange(len(df_sorted)), cum_perc, color='red', marker='o', label='Tích lũy (%)')
        ax2.set_ylabel('Tỷ lệ tích lũy (%)')
        ax1.set_title('Biểu đồ Pareto: Doanh số & tỷ trọng tích lũy')
        fig3.tight_layout()
        st.pyplot(fig3)
    except Exception as e:
        st.error(f"Lỗi khi vẽ Pareto chart: {e}")

elif chart_type == "Pie":
    try:
        fig4, ax4 = plt.subplots(figsize=(6,6))
        s = df.groupby('Nhóm khách hàng')['Tổng bán trừ trả hàng'].sum()
        ax4.pie(s, labels=s.index, autopct='%1.1f%%')
        ax4.set_title('Tỷ trọng doanh số theo nhóm khách hàng')
        st.pyplot(fig4)
    except Exception as e:
        st.error(f"Lỗi khi vẽ Pie chart: {e}")

# ===== XUẤT FILE ĐẸP, TẢI VỀ =====
st.subheader("4. Tải file kết quả định dạng màu nhóm")

output_file = 'sales_report_dep.xlsx'
df.to_excel(output_file, index=False)

wb = load_workbook(output_file)
ws = wb.active

header_fill = PatternFill(start_color='FFE699', end_color='FFE699', fill_type='solid')
header_font = Font(bold=True, color='000000')
header_align = Alignment(horizontal='center', vertical='center')
for col in range(1, ws.max_column + 1):
    cell = ws.cell(row=1, column=col)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = header_align

money_keywords = ['bán', 'doanh số', 'tiền', 'hoa hồng', 'comm', 'VND']
cols_money = [col[0].column for col in ws.iter_cols(1, ws.max_column)
              if any(key in (col[0].value or '').lower() for key in money_keywords)]

col_makh = [cell.value for cell in ws[1]].index('Mã khách hàng')+1
col_role = [cell.value for cell in ws[1]].index('Nhóm khách hàng')+1

all_codes = [str(ws.cell(row=i, column=col_makh).value) for i in range(2, ws.max_row+1)]
prefix_groups = {}
for length in range(len(max(all_codes, key=len)), 0, -1):
    prefix_count = {}
    for code in all_codes:
        if len(code) < length:
            continue
        prefix = code[:length]
        prefix_count.setdefault(prefix, []).append(code)
    for prefix, codes in prefix_count.items():
        if len(codes) > 1:
            prefix_groups[prefix] = codes

row_to_prefix = {}
for idx, code in enumerate(all_codes):
    best_prefix = ''
    best_len = 0
    for prefix in prefix_groups.keys():
        if code.startswith(prefix) and len(prefix) > best_len:
            best_prefix = prefix
            best_len = len(prefix)
    row_to_prefix[idx+2] = best_prefix if best_prefix else code

prefix_set = set(row_to_prefix.values())
prefix_list = sorted(prefix_set)
def get_contrasting_color(idx, total):
    h = idx / total
    r, g, b = colorsys.hsv_to_rgb(h, 0.65, 1)
    return "%02X%02X%02X" % (int(r*255), int(g*255), int(b*255))
prefix_to_color = {prefix: PatternFill(start_color=get_contrasting_color(i, len(prefix_list)),
                                       end_color=get_contrasting_color(i, len(prefix_list)),
                                       fill_type='solid')
                   for i, prefix in enumerate(prefix_list)}

for row in range(2, ws.max_row + 1):
    role = ws.cell(row=row, column=col_role).value
    if role == 'Trailblazer':
        fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    else:
        fill = prefix_to_color[row_to_prefix[row]]
    for col in range(1, ws.max_column + 1):
        ws.cell(row=row, column=col).fill = fill

for col in range(1, ws.max_column + 1):
    for row in range(2, ws.max_row+1):
        cell = ws.cell(row=row, column=col)
        if col in cols_money:
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal='right', vertical='center')
        else:
            cell.alignment = Alignment(horizontal='center', vertical='center')

for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        val = str(cell.value) if cell.value else ""
        max_length = max(max_length, len(val.encode('utf8'))//2+2)
    ws.column_dimensions[column].width = max(10, min(40, max_length))

bio = BytesIO()
wb.save(bio)
st.download_button(
    label="📥 Tải file Excel đã định dạng",
    data=bio.getvalue(),
    file_name=output_file,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ===== Footer =====
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown(
    "<div style='text-align:center;font-size:16px;color:#1570af;font-weight:600;'>Hotline: 0909.625.808</div>",
    unsafe_allow_html=True)
st.markdown(
    "<div style='text-align:center;font-size:14px;color:#555;'>Địa chỉ: Lầu 9, Pearl Plaza, 561A Điện Biên Phủ, P.25, Q. Bình Thạnh, TP.HCM</div>",
    unsafe_allow_html=True)
