import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objects as go

st.set_page_config(page_title="Phân tích phí & doanh thu", layout="wide")

st.title("Phân tích các khoản phí trên tổng doanh thu sản phẩm")

uploaded_file = st.file_uploader("Tải lên file báo cáo Excel", type=["xlsx"])

if uploaded_file:
    sheet = "Table 4"
    df = pd.read_excel(uploaded_file, sheet_name=sheet)
    st.subheader("Dữ liệu gốc (Table 4)")
    st.dataframe(df)

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

    revenue = float(df[df['Tổng kết thanh toán đã chuyển'] == "Tổng tiền sản phẩm"]['Số tiền (VND)'].values[0])
    total_fees = sum(fee_values)

    st.markdown("### Số liệu tổng hợp")
    st.write(f"**Tổng phí giao dịch:** {total_fees:,.0f} VND")
    st.write(f"**T
