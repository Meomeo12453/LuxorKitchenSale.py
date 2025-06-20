import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

st.set_page_config(layout="wide")
st.title("Sales Dashboard MiniApp")

# 1. Upload file Excel
uploaded_file = st.file_uploader("Upload file Excel dữ liệu bán hàng (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ Đã đọc file dữ liệu!")
    except Exception as e:
        st.error(f"❌ Lỗi khi đọc Excel: {e}")
        st.stop()
else:
    st.info("Vui lòng upload file Excel để bắt đầu.")
    st.stop()

# 2. Thiết lập lựa chọn hiển thị
chart_to_show = st.selectbox("Chọn biểu đồ muốn xem", ("All", "Bar", "Box", "Pie"))

# 3. Định nghĩa cấu trúc hoa hồng & tính toán override_sales
network = {
    'Catalyst':     {'comm_rate': 0.35, 'override_rate': 0.00},
    'Visionary':    {'comm_rate': 0.40, 'override_rate': 0.05},
    'Trailblazer':  {'comm_rate': 0.40, 'override_rate': 0.05},
}

def calculate_override(df_in):
    df2 = df_in.copy()
    df2['override_sales'] = 0
    for role in ['Catalyst', 'Visionary', 'Trailblazer']:
        staff = df2[df2['Role'] == role]
        for _, row in staff.iterrows():
            mask = df2['SuperiorCode'] == row['CustomerCode']
            subtotal = df2.loc[mask, ['Sales', 'override_sales']].sum().sum()
            df2.loc[df2['CustomerCode'] == row['CustomerCode'], 'override_sales'] = subtotal
    return df2

df = calculate_override(df)
df['comm_rate']     = df['Role'].map(lambda r: network[r]['comm_rate'])
df['override_rate'] = df['Role'].map(lambda r: network[r]['override_rate'])
df['override_comm'] = df['override_sales'] * df['override_rate']

st.subheader("Dữ liệu mẫu (5 dòng đầu):")
st.dataframe(df.head())

# 4. Vẽ biểu đồ theo lựa chọn
def plot_bar():
    fig, ax = plt.subplots(figsize=(8,4))
    ax.bar(df['CustomerCode'], df['Sales'])
    ax.set_title('Doanh số theo nhân viên')
    plt.xticks(rotation=45)
    st.pyplot(fig)

def plot_box():
    fig, ax = plt.subplots(figsize=(8,4))
    data = [df[df['Role']==r]['Sales'] for r in network]
    ax.boxplot(data, labels=list(network))
    ax.set_title('Phân phối doanh số theo Role')
    st.pyplot(fig)

def plot_pie():
    fig, ax = plt.subplots(figsize=(6,6))
    s = df.groupby('Role')['Sales'].sum()
    ax.pie(s, labels=s.index, autopct='%1.1f%%')
    ax.set_title('Tỷ trọng doanh số theo Role')
    ax.axis('equal')
    st.pyplot(fig)

if chart_to_show in ("All","Bar"):
    st.subheader("Biểu đồ Bar: Doanh số theo nhân viên")
    plot_bar()

if chart_to_show in ("All","Box"):
    st.subheader("Biểu đồ Box: Phân phối doanh số theo Role")
    plot_box()

if chart_to_show in ("All","Pie"):
    st.subheader("Biểu đồ Pie: Tỷ trọng doanh số theo Role")
    plot_pie()

# 5. Xuất báo cáo ra Excel và cho phép tải về
output = io.BytesIO()
df.to_excel(output, index=False, engine='openpyxl')
st.download_button("Tải báo cáo kết quả (Excel)", output.getvalue(), file_name="sales_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

