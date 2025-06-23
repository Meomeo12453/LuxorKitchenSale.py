import streamlit as st
from PIL import Image
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font
import colorsys
import os
import random

# ======= CSS ĐƠN GIẢN ĐẢM BẢO HIỂN THỊ =======
st.set_page_config(page_title="Sales Dashboard MiniApp", layout="wide")
st.markdown("""
    <style>
        .block-container {padding-top:0.9rem;}
        .stApp {background: #F7F8FA;}
        img {border-radius:0 !important;}
    </style>
""", unsafe_allow_html=True)

# ======= HIỂN THỊ LOGO, ĐẢM BẢO LUÔN CÓ FILE LOGO =======
logo_path = "logo-daba.png"  # Đổi tên file logo đúng
if os.path.exists(logo_path):
    img = Image.open(logo_path)
    desired_height = 38  # pixels, sửa nhỏ/lớn tại đây
    w, h = img.size
    new_width = int((w / h) * desired_height)
    img = img.resize((new_width, desired_height))
    st.markdown("<div style='display:flex;justify-content:center;'><img src='data:image/png;base64," +
        BytesIO(img.tobytes()).getvalue().hex() + "' width='" + str(new_width) +
        "' height='" + str(desired_height) + "'/></div>", unsafe_allow_html=True)
    st.image(img)
else:
    st.warning("❗️Không tìm thấy file logo-daba.png. Hãy upload file logo này vào thư mục cùng file app!")

st.markdown("<div style='text-align:center;font-size:16px;color:#1570af;font-weight:600;'>Hotline: 0909.625.808</div>", unsafe_allow_html=True)
st.markdown("<div style='text-align:center;font-size:14px;color:#555;'>Địa chỉ: Lầu 9, Pearl Plaza, 561A Điện Biên Phủ, P.25, Q. Bình Thạnh, TP.HCM</div>", unsafe_allow_html=True)
st.markdown("<hr style='margin:8px 0 18px 0;border:1px solid #EEE;'>", unsafe_allow_html=True)

# ======= CONTROL: BIỂU ĐỒ, LỌC NHÓM =======
st.markdown("### Tùy chọn phân tích")
col1, col2 = st.columns([2, 1])
with col1:
    chart_type = st.radio(
        "Chọn loại biểu đồ:",
        ["Biểu đồ cột chồng", "Sơ đồ Sunburst", "Biểu đồ Pareto", "Biểu đồ tròn (Pie)"],
        horizontal=True
    )
with col2:
    filter_nganh = st.multiselect("Lọc theo nhóm khách hàng:", ["Catalyst", "Visionary", "Trailblazer"], default=[])

st.markdown("<hr style='margin:10px 0 18px 0;border:1px solid #EEE;'>", unsafe_allow_html=True)

# ======= UPLOAD NHIỀU FILE EXCEL =======
st.markdown("### 1. Tải lên (tối đa 10) file Excel (.xlsx)")
uploaded_files = st.file_uploader("Chọn nhiều file cùng lúc (hoặc kéo thả)", type="xlsx", accept_multiple_files=True)
if not uploaded_files:
    st.info("💡 Vui lòng upload file Excel mẫu (có cột: Mã khách hàng, Tên khách hàng, Nhóm khách hàng, Tổng bán trừ trả hàng, Ghi chú)")
    st.stop()

if len(uploaded_files) > 10:
    st.warning("Chỉ upload tối đa 10 file!")
    st.stop()

# ======= ĐỌC & GHÉP FILE =======
dfs = []
for file in uploaded_files:
    try:
        df = pd.read_excel(file)
        # Chuẩn hóa cột
        df.columns = [str(col).strip() for col in df.columns]
        # Bắt buộc phải có các cột sau
        required_cols = ['Mã khách hàng', 'Tên khách hàng', 'Nhóm khách hàng', 'Tổng bán trừ trả hàng', 'Ghi chú']
        for col in required_cols:
            if col not in df.columns:
                st.error(f"File {file.name} thiếu cột bắt buộc: {col}")
                st.stop()
        dfs.append(df)
    except Exception as e:
        st.error(f"File {file.name} không đọc được: {e}")
        st.stop()

df = pd.concat(dfs, ignore_index=True)
df['Mã khách hàng'] = df['Mã khách hàng'].astype(str).str.strip()
df['Ghi chú'] = df['Ghi chú'].astype(str).str.strip().replace({'None': None, 'nan': None, 'NaN': None, '': None})
df['Tổng bán trừ trả hàng'] = pd.to_numeric(df['Tổng bán trừ trả hàng'], errors='coerce').fillna(0)

# ======= XÁC ĐỊNH QUAN HỆ HỆ THỐNG THEO "GHI CHÚ" =======
all_codes = set(df['Mã khách hàng'])
def get_parent_id(x):
    if pd.isnull(x) or x is None: return None
    x = str(x).strip()
    return x if x in all_codes else None
df['parent_id'] = df['Ghi chú'].apply(get_parent_id)

# Xây parent_map cho hệ thống đa tầng
parent_map = {}
for idx, row in df.iterrows():
    pid = row['parent_id']
    code = row['Mã khách hàng']
    if pd.notnull(pid) and pid is not None:
        parent_map.setdefault(str(pid), []).append(str(code))

def get_all_descendants(code, parent_map):
    result = []
    direct = parent_map.get(str(code), [])
    result.extend(direct)
    for child in direct:
        result.extend(get_all_descendants(child, parent_map))
    return result

desc_counts = []
ds_he_thong = []
for idx, row in df.iterrows():
    code = str(row['Mã khách hàng'])
    descendants = get_all_descendants(code, parent_map)
    desc_counts.append(len(descendants))
    doanhso = df[df['Mã khách hàng'].isin(descendants)]['Tổng bán trừ trả hàng'].sum() if descendants else 0
    ds_he_thong.append(doanhso)
df['Số cấp dưới'] = desc_counts
df['Doanh số hệ thống'] = ds_he_thong

# ======= HOA HỒNG =======
network = {
    'Catalyst':     {'comm_rate': 0.35, 'override_rate': 0.00},
    'Visionary':    {'comm_rate': 0.40, 'override_rate': 0.05},
    'Trailblazer':  {'comm_rate': 0.40, 'override_rate': 0.05},
}
df['comm_rate']     = df['Nhóm khách hàng'].map(lambda r: network.get(r, {}).get('comm_rate', 0))
df['override_rate'] = df['Nhóm khách hàng'].map(lambda r: network.get(r, {}).get('override_rate', 0))
df['override_comm'] = df['Doanh số hệ thống'] * df['override_rate']

if filter_nganh:
    df = df[df['Nhóm khách hàng'].isin(filter_nganh)]

st.markdown("### 2. Bảng dữ liệu đại lý đã xử lý")
st.dataframe(df, use_container_width=True, hide_index=True)

# ======= BIỂU ĐỒ PHÂN TÍCH DỮ LIỆU =======
st.markdown("### 3. Biểu đồ phân tích dữ liệu")
import matplotlib.pyplot as plt
import plotly.express as px
import numpy as np

if chart_type == "Biểu đồ cột chồng":
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

elif chart_type == "Sơ đồ Sunburst":
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

elif chart_type == "Biểu đồ Pareto":
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

elif chart_type == "Biểu đồ tròn (Pie)":
    try:
        fig4, ax4 = plt.subplots(figsize=(6,6))
        s = df.groupby('Nhóm khách hàng')['Tổng bán trừ trả hàng'].sum()
        ax4.pie(s, labels=s.index, autopct='%1.1f%%')
        ax4.set_title('Tỷ trọng doanh số theo nhóm khách hàng')
        st.pyplot(fig4)
    except Exception as e:
        st.error(f"Lỗi khi vẽ Pie chart: {e}")

st.markdown("<hr style='margin:10px 0 18px 0;border:1px solid #EEE;'>", unsafe_allow_html=True)

# ======= EXPORT FILE MÀU NHÓM (NẾU CẦN) =======
output_file = 'sales_report_dep.xlsx'
df.to_excel(output_file, index=False)

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font

wb = load_workbook(output_file)
ws = wb.active

# Tô màu nhóm theo parent_id (cấp hệ thống)
def pastel_color(seed_val):
    random.seed(str(seed_val))
    h = random.random()
    s = 0.27 + random.random()*0.09
    v = 0.96
    r, g, b = colorsys.hsv_to_rgb(h, s, v)
    return "%02X%02X%02X" % (int(r*255), int(g*255), int(b*255))

col_makh = [cell.value for cell in ws[1]].index('Mã khách hàng')+1
col_parent = [cell.value for cell in ws[1]].index('parent_id')+1
ma_cha_list = df[df['Mã khách hàng'].isin(df['parent_id'].dropna())]['Mã khách hàng'].unique().tolist()
ma_cha_to_color = {ma_cha: PatternFill(start_color=pastel_color(ma_cha), end_color=pastel_color(ma_cha), fill_type='solid') for ma_cha in ma_cha_list}

for row in range(2, ws.max_row + 1):
    ma_kh = str(ws.cell(row=row, column=col_makh).value)
    parent_id = ws.cell(row=row, column=col_parent).value
    if ma_kh in ma_cha_to_color:
        fill = ma_cha_to_color[ma_kh]
    elif parent_id in ma_cha_to_color:
        fill = ma_cha_to_color[parent_id]
    else:
        fill = PatternFill(fill_type=None)
    for col in range(1, ws.max_column + 1):
        ws.cell(row=row, column=col).fill = fill

header_fill = PatternFill(start_color='FFE699', end_color='FFE699', fill_type='solid')
header_font = Font(bold=True, color='000000')
header_align = Alignment(horizontal='center', vertical='center')
for col in range(1, ws.max_column + 1):
    cell = ws.cell(row=1, column=col)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = header_align

bio = BytesIO()
wb.save(bio)
st.download_button(
    label="📥 Tải file Excel đã định dạng",
    data=bio.getvalue(),
    file_name=output_file,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ======= FOOTER =======
st.markdown("<hr style='margin:10px 0 18px 0;border:1px solid #EEE;'>", unsafe_allow_html=True)
st.markdown("<div style='text-align:center;font-size:16px;color:#1570af;font-weight:600;'>Hotline: 0909.625.808</div>", unsafe_allow_html=True)
st.markdown("<div style='text-align:center;font-size:14px;color:#555;'>Địa chỉ: Lầu 9, Pearl Plaza, 561A Điện Biên Phủ, P.25, Q. Bình Thạnh, TP.HCM</div>", unsafe_allow_html=True)
