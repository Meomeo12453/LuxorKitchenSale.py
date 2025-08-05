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
import random
import base64
import uuid
from matplotlib.backends.backend_pdf import PdfPages

st.set_page_config(page_title="Sales Dashboard MiniApp", layout="wide")
for _ in range(4):
    st.write("")

st.markdown("""
    <style>
    .block-container {padding-top:0.7rem; max-width:100vw !important;}
    .stApp {background: #F7F8FA;}
    img { border-radius: 0 !important; }
    h1, h2, h3 { font-size: 1.18rem !important; font-weight:600; }
    </style>
""", unsafe_allow_html=True)

LOGO_PATHS = [
    "logo-daba.png",
    "ef5ac011-857d-4b32-bd70-ef9ac3817106.png"
]
logo = None
for path in LOGO_PATHS:
    if os.path.exists(path):
        logo = Image.open(path)
        break

if logo is not None:
    desired_height = 36
    w, h = logo.size
    new_width = int((w / h) * desired_height)
    logo_resized = logo.resize((new_width, desired_height))
    buffered = BytesIO()
    logo_resized.save(buffered, format="PNG")
    img_str = base64.b64encode(buffered.getvalue()).decode()
    st.markdown(
        f"""
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;width:100%;padding-top:4px;padding-bottom:0;">
            <img src="data:image/png;base64,{img_str}" 
                 width="{new_width}" height="{desired_height}" style="display:block;margin:auto;" />
            <div style="height:5px;"></div>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown(
    "<div style='text-align:center;font-size:20px;color:#1570af;font-weight:600;'>BẢNG TÍNH HOA HỒNG CÔNG TY TNHH DABA SAIGON</div>",
    unsafe_allow_html=True)
st.markdown(
    "<div style='text-align:center;font-size:14px;color:#555;'>Hotline 0909.625.808 Địa chỉ: Lầu 9, Pearl Plaza, 561A Điện Biên Phủ, P.25, Q. Bình Thạnh, TP.HCM</div>",
    unsafe_allow_html=True)
st.markdown("<hr style='margin:10px 0 20px 0;border:1px solid #EEE;'>", unsafe_allow_html=True)

# ========== CONTROL ==========
st.markdown("### 🔎 Tùy chọn phân tích")
col1, col2 = st.columns([2, 1])
with col1:
    chart_type = st.radio(
        "Chọn loại biểu đồ:",
        ["Biểu đồ cột chồng", "Sơ đồ Sunburst", "Biểu đồ Pareto", "Biểu đồ tròn (Pie)"],
        horizontal=True
    )
with col2:
    filter_nganh = st.multiselect("Lọc theo nhóm khách hàng:", ["Catalyst", "Visionary", "Trailblazer"], default=[])

st.markdown("<hr style='margin:10px 0 20px 0;border:1px solid #EEE;'>", unsafe_allow_html=True)

# ======= MULTI FILE UPLOAD =======
st.markdown("### 1. Tải lên tối đa 10 file Excel (.xlsx)")
uploaded_files = st.file_uploader(
    "**Chọn nhiều file hoặc kéo thả nhiều file Excel**",
    type="xlsx",
    accept_multiple_files=True,
    help="Chỉ nhận Excel, <200MB mỗi file. Các file phải cùng cấu trúc cột."
)
if not uploaded_files:
    st.info("💡 Hãy upload 1 hoặc nhiều file Excel mẫu để bắt đầu sử dụng Dashboard.")
    with st.expander("📋 Xem hướng dẫn & file mẫu", expanded=False):
        st.markdown(
            "- Chọn hoặc kéo thả **1–10 file Excel**.\n"
            "- File cần các cột: **Mã khách hàng, Tên khách hàng, Nhóm khách hàng, Tổng bán trừ trả hàng, Ghi chú**.\n"
            "- Nếu lỗi, kiểm tra lại tiêu đề cột trong file Excel."
        )
    st.stop()

# ===== GỘP & LÀM SẠCH DỮ LIỆU =====
dfs = []
for f in uploaded_files[:10]:
    dft = pd.read_excel(f)
    dfs.append(dft)
df = pd.concat(dfs, ignore_index=True)

# ===== CẢNH BÁO TRÙNG MÃ KHÁCH HÀNG =====
duplicated_mask = df.duplicated(subset=['Mã khách hàng'], keep=False)
if duplicated_mask.any():
    dup_kh = df.loc[duplicated_mask, 'Mã khách hàng']
    dup_kh_list = dup_kh.value_counts().index.tolist()
    dup_kh_str = ", ".join(str(x) for x in dup_kh_list)
    st.warning(
        f"⚠️ Có {len(dup_kh_list)} mã khách hàng bị trùng trong file dữ liệu: **{dup_kh_str}**. "
        "Chỉ giữ lại dòng đầu tiên cho mỗi mã. Vui lòng kiểm tra lại file gốc!"
    )

df['Mã khách hàng'] = df['Mã khách hàng'].astype(str).str.strip()
df['Ghi chú'] = df['Ghi chú'].astype(str).str.strip()
df['Ghi chú'] = df['Ghi chú'].replace({'None': None, 'nan': None, 'NaN': None, '': None})
df['Tổng bán trừ trả hàng'] = pd.to_numeric(df['Tổng bán trừ trả hàng'], errors='coerce').fillna(0)
df = df.drop_duplicates(subset=['Mã khách hàng'], keep='first')

all_codes = set(df['Mã khách hàng'])

def get_parent_id(x):
    if pd.isnull(x) or x is None:
        return None
    x = str(x).strip()
    return x if x in all_codes else None
df['parent_id'] = df['Ghi chú'].apply(get_parent_id)

parent_map = {}
for idx, row in df.iterrows():
    pid = row['parent_id']
    code = row['Mã khách hàng']
    if pd.notnull(pid) and pid is not None:
        parent_map.setdefault(pid, []).append(code)

def get_all_descendants(code, parent_map, visited=None):
    if visited is None:
        visited = set()
    result = []
    children = parent_map.get(code, [])
    for child in children:
        if child not in visited:
            visited.add(child)
            result.append(child)
            result.extend(get_all_descendants(child, parent_map, visited))
    return result

desc_counts = []
ds_he_thong = []
for idx, row in df.iterrows():
    code = row['Mã khách hàng']
    descendants = get_all_descendants(code, parent_map, visited=set([code]))
    desc_counts.append(len(descendants))
    doanhso = df[df['Mã khách hàng'].isin(descendants)]['Tổng bán trừ trả hàng'].sum() if descendants else 0
    ds_he_thong.append(doanhso)
df['Số cấp dưới'] = desc_counts
df['Doanh số hệ thống'] = ds_he_thong

network = {
    'Catalyst':     {'comm_rate': 0.35, 'override_rate': 0.00},
    'Visionary':    {'comm_rate': 0.40, 'override_rate': 0.05},
    'Trailblazer':  {'comm_rate': 0.40, 'override_rate': 0.05},
}
df['comm_rate']     = df['Nhóm khách hàng'].map(lambda r: network.get(r, {}).get('comm_rate', 0))
df['override_rate'] = df['Nhóm khách hàng'].map(lambda r: network.get(r, {}).get('override_rate', 0))

# ==== TÍNH VƯỢT CẤP & GẮN VÀO DATAFRAME ====
trailblazer_codes = df[df['Nhóm khách hàng'] == 'Trailblazer']['Mã khách hàng'].astype(str)
catalyst_children = df[(df['Nhóm khách hàng'] == 'Catalyst') & (df['parent_id'].notnull())]
catalyst_children = catalyst_children[catalyst_children['parent_id'].isin(trailblazer_codes)]
vuot_cap_ds = catalyst_children.groupby('parent_id')['Tổng bán trừ trả hàng'].sum()
vuot_cap_hh = vuot_cap_ds * 0.10
df['Doanh số vượt cấp'] = df['Mã khách hàng'].astype(str).map(vuot_cap_ds).fillna(0)
df['Hoa hồng vượt cấp'] = df['Mã khách hàng'].astype(str).map(vuot_cap_hh).fillna(0)
catalyst_sys_map = catalyst_children.set_index('Mã khách hàng')['parent_id'].to_dict()
df['vuot_cap_trailblazer'] = df['Mã khách hàng'].map(catalyst_sys_map)

# === ĐIỀU CHỈNH DOANH SỐ HỆ THỐNG CHO TRAILBLAZER ===
tb_catalyst_vuotcap_doanhso = vuot_cap_ds.to_dict()
for idx, row in df.iterrows():
    if row['Nhóm khách hàng'] == 'Trailblazer':
        minus = tb_catalyst_vuotcap_doanhso.get(row['Mã khách hàng'], 0)
        df.at[idx, 'Doanh số hệ thống'] = max(row['Doanh số hệ thống'] - minus, 0)

# ==== TÍNH override_comm SAU khi cập nhật Doanh số hệ thống! ====
df['override_comm'] = df['Doanh số hệ thống'] * df['override_rate']

# Sắp xếp lại thứ tự cột nếu cần
cols = list(df.columns)
if 'Hoa hồng vượt cấp' in cols and 'Doanh số vượt cấp' in cols:
    cols.remove('Doanh số vượt cấp')
    idx_hhvc = cols.index('Hoa hồng vượt cấp')
    cols.insert(idx_hhvc, 'Doanh số vượt cấp')
df = df[cols]

if filter_nganh:
    df = df[df['Nhóm khách hàng'].isin(filter_nganh)]

st.markdown("### 2. Bảng dữ liệu đại lý đã xử lý")
st.dataframe(df, use_container_width=True, hide_index=True)

# ====== Tạo các biểu đồ (matplotlib) và lưu thành các biến fig ======
fig1, ax1 = plt.subplots(figsize=(12,5))
ind = np.arange(len(df))
ax1.bar(ind, df['Tổng bán trừ trả hàng'], width=0.5, label='Tổng bán cá nhân')
ax1.bar(ind, df['override_comm'], width=0.5, bottom=df['Tổng bán trừ trả hàng'], label='Hoa hồng hệ thống')
ax1.set_ylabel('Số tiền (VND)')
ax1.set_title('Tổng bán & Hoa hồng hệ thống từng cá nhân')
ax1.set_xticks(ind)
ax1.set_xticklabels(df['Tên khách hàng'], rotation=60, ha='right')
ax1.legend()

fig2, ax2 = plt.subplots(figsize=(10,5))
df_sorted = df.sort_values('Tổng bán trừ trả hàng', ascending=False)
cum_sum = df_sorted['Tổng bán trừ trả hàng'].cumsum()
cum_perc = 100 * cum_sum / df_sorted['Tổng bán trừ trả hàng'].sum()
ax2.bar(np.arange(len(df_sorted)), df_sorted['Tổng bán trừ trả hàng'], label="Doanh số")
ax2.set_ylabel('Doanh số')
ax2.set_xticks(range(len(df_sorted)))
ax2.set_xticklabels(df_sorted['Tên khách hàng'], rotation=60, ha='right')
ax2_2 = ax2.twinx()
ax2_2.plot(np.arange(len(df_sorted)), cum_perc, color='red', marker='o', label='Tích lũy (%)')
ax2_2.set_ylabel('Tỷ lệ tích lũy (%)')
ax2.set_title('Biểu đồ Pareto: Doanh số & tỷ trọng tích lũy')
fig2.tight_layout()

fig3, ax3 = plt.subplots(figsize=(6,6))
s = df.groupby('Nhóm khách hàng')['Tổng bán trừ trả hàng'].sum()
ax3.pie(s, labels=s.index, autopct='%1.1f%%')
ax3.set_title('Tỷ trọng doanh số theo nhóm khách hàng')

st.markdown("### 3. Biểu đồ phân tích dữ liệu")
st.pyplot(fig1)
st.pyplot(fig2)
st.pyplot(fig3)

pdf_bytes = BytesIO()
with PdfPages(pdf_bytes) as pdf:
    for fig in [fig1, fig2, fig3]:
        pdf.savefig(fig, bbox_inches='tight')
pdf_bytes.seek(0)
st.download_button(
    "📥 Tải tất cả biểu đồ thành 1 file PDF",
    data=pdf_bytes.getvalue(),
    file_name="all_charts.pdf",
    mime="application/pdf"
)

st.markdown("### 4. Tải file kết quả định dạng màu vượt cấp & cha–con")

output_file = f'sales_report_dep_{uuid.uuid4().hex[:6]}.xlsx'
df_export = df.sort_values(by=['parent_id', 'Mã khách hàng'], ascending=[True, True], na_position='last')
df_export.to_excel(output_file, index=False)

wb = load_workbook(output_file)
ws = wb.active
col_names = [cell.value for cell in ws[1]]
col_makh = col_names.index('Mã khách hàng')+1
col_parent = col_names.index('parent_id')+1 if 'parent_id' in col_names else None
col_vuotcap = col_names.index('vuot_cap_trailblazer')+1 if 'vuot_cap_trailblazer' in col_names else None

def pastel_color(seed_val):
    random.seed(str(seed_val))
    h = random.random()
    s = 0.28 + random.random()*0.09
    v = 0.97
    r, g, b = colorsys.hsv_to_rgb(h, s, v)
    return "%02X%02X%02X" % (int(r*255), int(g*255), int(b*255))

trailblazer_vuotcap = set(df['vuot_cap_trailblazer'].dropna().unique()).union(df[df['Nhóm khách hàng']=='Trailblazer']['Mã khách hàng'])
trailblazer_to_color = {tb: PatternFill(start_color=pastel_color(tb+"vuotcap"), end_color=pastel_color(tb+"vuotcap"), fill_type='solid') for tb in trailblazer_vuotcap}

ma_cha_list = df_export[df_export['Mã khách hàng'].isin(df_export['parent_id'].dropna())]['Mã khách hàng'].unique().tolist() if col_parent else []
ma_cha_to_color = {ma_cha: PatternFill(start_color=pastel_color(ma_cha), end_color=pastel_color(ma_cha), fill_type='solid') for ma_cha in ma_cha_list}

for row in range(2, ws.max_row + 1):
    ma_kh = str(ws.cell(row=row, column=col_makh).value)
    parent_id = ws.cell(row=row, column=col_parent).value if col_parent else None
    vuotcap_tb = ws.cell(row=row, column=col_vuotcap).value if col_vuotcap else None
    if (vuotcap_tb and vuotcap_tb in trailblazer_to_color):
        fill = trailblazer_to_color[vuotcap_tb]
    elif ma_kh in trailblazer_to_color:
        fill = trailblazer_to_color[ma_kh]
    elif col_parent and ma_kh in ma_cha_to_color:
        fill = ma_cha_to_color[ma_kh]
    elif col_parent and parent_id in ma_cha_to_color:
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

cols_to_drop = [
    "vuot_cap_trailblazer", "Loại khách", "Chi nhánh tạo", "Khu vực giao hàng", "Phường/Xã", "Số CMND/CCCD",
    "Ngày sinh", "Giới tính", "Email", "Facebook", "parent_id", "Người tạo", "Ngày tạo", "Tổng bán", "Trạng thái"
]
ws_header = [cell.value for cell in ws[1]]
for col_name in cols_to_drop:
    if col_name in ws_header:
        col_idx = ws_header.index(col_name) + 1
        ws.delete_cols(col_idx)
        ws_header.pop(col_idx - 1)

header_map = {
    "comm_rate": "Chiet_khau",
    "override_rate": "TL_Hoa_Hong",
    "override_comm": "Hoa_hong_he_thong"
}
for col in range(1, ws.max_column + 1):
    cell = ws.cell(row=1, column=col)
    if cell.value in header_map:
        cell.value = header_map[cell.value]

bio = BytesIO()
wb.save(bio)
downloaded = st.download_button(
    label="📥 Tải file Excel đã định dạng",
    data=bio.getvalue(),
    file_name=output_file,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
if downloaded:
    st.toast("✅ Đã tải xuống!", icon="✅")
