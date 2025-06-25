import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Biểu đồ phí giao dịch vs doanh thu", layout="centered")

st.title("Biểu đồ phí giao dịch và doanh thu sản phẩm")
uploaded_file = st.file_uploader("Tải lên file báo cáo Excel", type=["xlsx"])

if uploaded_file:
    sheet = "Table 4"
    df = pd.read_excel(uploaded_file, sheet_name=sheet)
    
    st.subheader("Dữ liệu sheet Table 4")
    st.dataframe(df.head(15))

    # Lọc các mục phí giao dịch
    fee_items = [
        'Phí cố định',
        'Phí thanh toán',
        'Phí hoa hồng Tiếp thị liên kết',
        'Phí dịch vụ PiShip'
    ]
    df_fees = df[df['Tổng kết thanh toán đã chuyển'].astype(str).str.contains("Phí", na=False)]
    df_fees = df_fees.set_index('Tổng kết thanh toán đã chuyển')
    fee_values = []
    for item in fee_items:
        try:
            val = float(df_fees.loc[item, 'Số tiền (VND)'])
            fee_values.append(abs(val))
        except Exception:
            fee_values.append(0)

    # Tổng doanh thu sản phẩm
    try:
        revenue = float(df[df['Tổng kết thanh toán đã chuyển'] == "Tổng tiền sản phẩm"]['Số tiền (VND)'].values[0])
    except:
        st.error("Không tìm thấy dòng 'Tổng tiền sản phẩm'. Hãy kiểm tra lại file dữ liệu!")
        st.stop()
    total_fees = sum(fee_values)

    # Biểu đồ phần trăm phí giao dịch
    st.subheader("Tỷ lệ phần trăm từng mục phí giao dịch so với tổng doanh thu sản phẩm")
    percentages = [v / revenue * 100 for v in fee_values]
    fig, ax = plt.subplots()
    ax.bar(fee_items, percentages, color='orange')
    ax.set_ylabel("Phần trăm (%)")
    ax.set_title("Tỷ lệ từng mục phí giao dịch so với doanh thu")
    ax.set_ylim(0, max(percentages)*1.25)
    plt.xticks(rotation=15)
    st.pyplot(fig)

    # Biểu đồ tổng phí giao dịch vs doanh thu
    st.subheader("So sánh tổng phí giao dịch và tổng doanh thu sản phẩm")
    fig2, ax2 = plt.subplots()
    ax2.bar(['Tổng phí giao dịch', 'Tổng doanh thu sản phẩm'], [total_fees, revenue], color='orange')
    ax2.set_ylabel("Giá trị (VND)")
    ax2.set_title("Tổng phí giao dịch vs Tổng doanh thu sản phẩm")
    st.pyplot(fig2)

    # Số liệu tổng hợp
    st.write(f"**Tổng phí giao dịch:** {total_fees:,.0f} VND")
    st.write(f"**Tổng doanh thu sản phẩm:** {revenue:,.0f} VND")
    st.write(f"**Tỷ lệ tổng phí/doanh thu:** {total_fees / revenue * 100:.2f} %")

else:
    st.info("Vui lòng tải lên file Excel để xem biểu đồ.")
