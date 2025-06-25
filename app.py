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

    # ---- Xử lý phí giao dịch & doanh thu ----
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
    st.write(f"**Tổng doanh thu sản phẩm:** {revenue:,.0f} VND")
    st.write(f"**Tỷ lệ tổng phí/doanh thu:** {total_fees / revenue * 100:.2f} %")

    # ---- Biểu đồ 1: Bar chart phần trăm phí giao dịch ----
    st.subheader("Tỷ lệ phần trăm từng mục phí giao dịch")
    percentages = [v / revenue * 100 for v in fee_values]
    fig, ax = plt.subplots()
    ax.bar(fee_items, percentages, color='orange')
    ax.set_ylabel("Phần trăm (%)")
    ax.set_title("Từng mục phí giao dịch so với doanh thu")
    plt.xticks(rotation=15)
    st.pyplot(fig)

    # ---- Biểu đồ 2: Tổng phí vs tổng doanh thu ----
    st.subheader("So sánh tổng phí giao dịch và tổng doanh thu sản phẩm")
    fig2, ax2 = plt.subplots()
    ax2.bar(['Tổng phí giao dịch', 'Tổng doanh thu sản phẩm'], [total_fees, revenue], color='orange')
    ax2.set_ylabel("Giá trị (VND)")
    ax2.set_title("Tổng phí giao dịch vs Tổng doanh thu sản phẩm")
    st.pyplot(fig2)

    # ---- Biểu đồ 3: Pie chart các loại phí ----
    st.subheader("Tỷ trọng từng loại phí giao dịch (Pie chart)")
    fig3, ax3 = plt.subplots()
    ax3.pie(fee_values, labels=fee_items, autopct='%1.1f%%', startangle=90)
    ax3.axis('equal')
    st.pyplot(fig3)

    # ---- Biểu đồ 4: Waterfall chart (thác nước) ----
    st.subheader("Biểu đồ thác nước: Dòng tiền trên tổng doanh thu sản phẩm")
    items = [
        {"label": "Tổng tiền sản phẩm", "value": revenue},
        {"label": "Phí cố định", "value": -fee_values[0]},
        {"label": "Phí thanh toán", "value": -fee_values[1]},
        {"label": "Phí tiếp thị liên kết", "value": -fee_values[2]},
        {"label": "Phí dịch vụ PiShip", "value": -fee_values[3]},
        {"label": "Doanh thu sau phí", "value": revenue - total_fees}
    ]
    labels = [i["label"] for i in items]
    values = [i["value"] for i in items]
    measure = ["relative"] * (len(items)-2) + ["total","total"]

    fig4 = go.Figure(go.Waterfall(
        name="Dòng tiền",
        orientation="v",
        measure=["absolute"] + ["relative"] * (len(items)-2) + ["total"],
        x=labels,
        y=[revenue] + [-v for v in fee_values] + [revenue - total_fees],
        text=[f"{abs(v):,.0f}₫" for v in [revenue] + [-v for v in fee_values] + [revenue - total_fees]],
        connector={"line": {"color": "orange"}},
    ))
    fig4.update_layout(
        title="Biểu đồ thác nước: Dòng tiền trên tổng doanh thu sản phẩm",
        showlegend=False,
        height=500
    )
    st.plotly_chart(fig4, use_container_width=True)

else:
    st.info("Vui lòng tải lên file Excel để xem biểu đồ.")

st.caption("Made by ChatGPT – yêu cầu tuỳ chỉnh code vui lòng nhắn trực tiếp.")
