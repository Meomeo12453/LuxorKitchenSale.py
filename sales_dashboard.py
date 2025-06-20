import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("Sales Dashboard MiniApp")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ Đã đọc file!")
    
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
    df['comm_rate'] = df['Role'].map(lambda r: network[r]['comm_rate'])
    df['override_rate'] = df['Role'].map(lambda r: network[r]['override_rate'])
    df['override_comm'] = df['override_sales'] * df['override_rate']
    
    st.dataframe(df.head())

    # Charts
    chart = st.radio("Chọn biểu đồ", ("All", "Bar", "Box", "Pie"))
    if chart in ("All", "Bar"):
        st.subheader("Doanh số theo nhân viên")
        fig, ax = plt.subplots(figsize=(8,4))
        ax.bar(df['CustomerCode'], df['Sales'])
        ax.set_title('Doanh số theo nhân viên')
        plt.xticks(rotation=45)
        st.pyplot(fig)

    if chart in ("All", "Box"):
        st.subheader("Phân phối doanh số theo Role")
        fig, ax = plt.subplots(figsize=(8,4))
        data = [df[df['Role']==r]['Sales'] for r in network]
        ax.boxplot(data, labels=list(network))
        ax.set_title('Phân phối doanh số theo Role')
        st.pyplot(fig)

    if chart in ("All", "Pie"):
        st.subheader("Tỷ trọng doanh số theo Role")
        fig, ax = plt.subplots(figsize=(6,6))
        s = df.groupby('Role')['Sales'].sum()
        ax.pie(s, labels=s.index, autopct='%1.1f%%')
        ax.set_title('Tỷ trọng doanh số theo Role')
        ax.axis('equal')
        st.pyplot(fig)
    
    # Xuất báo cáo
    import io
    output = io.BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    st.download_button("Tải file kết quả Excel", output.getvalue(), file_name="sales_report.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
else:
    st.info("Vui lòng tải file Excel để bắt đầu")

