import streamlit as st
import pandas as pd
import pdfplumber
import re
import matplotlib.pyplot as plt
import plotly.graph_objects as go

st.set_page_config(page_title="Dashboard PDF", layout="wide")

st.title("Báo cáo PDF doanh thu & chi phí (chart trực quan & xuất file)")

uploaded_pdf = st.file_uploader("Tải lên file PDF báo cáo", type=["pdf"])

def parse_date_range(pdf_text):
    lines = pdf_text.split('\n')
    from_date = None
    to_date = None
    for line in lines:
        if "Báo cáo từ" in line:
            from_date_match = re.search(r"-?(\d{8})", line)
            if from_date_match:
                from_date = from_date_match.group(1)
        if "đến" in line.lower():
            to_date_match = re.search(r"-?(\d{8})", line)
            if to_date_match:
                to_date = to_date_match.group(1)
    return from_date, to_date

def parse_pdf_summary(pdf_text):
    pattern = r"(?P<item>[^\d\n]+)\s+(?P<amount>[-₫\d,.]+)"
    summary = re.findall(pattern, pdf_text)
    rows = []
    for item, amount in summary:
        item = item.strip().replace("\n", " ")
        amount = amount.strip().replace("₫", "").replace(",", "").replace(".", "")
        try:
            value = float(amount.replace("-", "")) * (-1 if "-" in amount else 1)
        except:
            continue
        if len(item) < 2 or not re.search(r'\d', amount):
            continue
        rows.append((item, value))
    return pd.DataFrame(rows, columns=["Hạng mục", "Giá trị (VND)"])

if uploaded_pdf:
    with pdfplumber.open(uploaded_pdf) as pdf:
        all_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"

    # Lấy ngày báo cáo từ file PDF
    from_date, to_date = parse_date_range(all_text)
    # Hiển thị thời gian báo cáo ở đầu dashboard
    if from_date and to_date:
        st.success(f"**Báo cáo từ ngày:** {from_date[:4]}-{from_date[4:6]}-{from_date[6:]}  &nbsp;&nbsp; **đến ngày:** {to_date[:4]}-{to_date[4:6]}-{to_date[6:]}")
    else:
        st.warning("Không nhận diện được thời gian báo cáo trong file PDF.")

    df_summary = parse_pdf_summary(all_text)
    st.subheader("Bảng tổng hợp chi phí & doanh thu")
    st.dataframe(df_summary)

    # Lấy tổng doanh thu
    revenue_row = df_summary[df_summary["Hạng mục"].str.contains("Tổng tiền sản phẩm", case=False)]
    if not revenue_row.empty:
        revenue = float(revenue_row["Giá trị (VND)"].iloc[0])
        df_summary["% trên doanh thu"] = df_summary["Giá trị (VND)"].apply(lambda x: x/revenue*100 if revenue else 0)
    else:
        st.warning("Không tìm thấy dòng 'Tổng tiền sản phẩm' trong bảng dữ liệu PDF.")
        revenue = None
        df_summary["% trên doanh thu"] = 0.0

    # 1. Bar chart: Chi phí vs % doanh thu
    st.subheader("Biểu đồ cột: Giá trị và phần trăm chi phí trên tổng doanh thu")
    fee_df = df_summary[(df_summary["Giá trị (VND)"]<0) & (df_summary["% trên doanh thu"]!=0)]
    fig, ax = plt.subplots(figsize=(10,5))
    bars = ax.bar(fee_df["Hạng mục"], abs(fee_df["Giá trị (VND)"]), color='orange')
    for bar, pct in zip(bars, abs(fee_df["% trên doanh thu"])):
        ax.annotate(f"{pct:.1f}%", xy=(bar.get_x() + bar.get_width()/2, bar.get_height()),
                    xytext=(0, 5), textcoords="offset points", ha="center", va="bottom", fontsize=10, color="black", fontweight='bold')
    ax.set_ylabel("Giá trị (VND)")
    ax.set_xlabel("Loại chi phí")
    ax.set_title("Giá trị & % từng loại chi phí so với tổng doanh thu")
    plt.xticks(rotation=18)
    st.pyplot(fig)

    # 2. Pie chart: Cơ cấu chi phí trên tổng phí
    st.subheader("Biểu đồ tròn (Pie): Cơ cấu từng loại chi phí")
    pie_labels = fee_df["Hạng mục"]
    pie_values = abs(fee_df["Giá trị (VND)"])
    fig2, ax2 = plt.subplots()
    wedges, texts, autotexts = ax2.pie(pie_values, labels=pie_labels, autopct='%1.1f%%', startangle=90)
    ax2.axis('equal')
    ax2.set_title("Tỷ trọng từng loại chi phí trong tổng phí")
    st.pyplot(fig2)

    # 3. Waterfall chart: Doanh thu - lần lượt trừ chi phí
    st.subheader("Biểu đồ thác nước: Quá trình trừ chi phí khỏi doanh thu")
    labels = ["Tổng doanh thu"] + fee_df["Hạng mục"].tolist() + ["Còn lại sau chi phí"]
    values = [revenue] + fee_df["Giá trị (VND)"].tolist() + [revenue + fee_df["Giá trị (VND)"].sum()]
    measure = ["absolute"] + ["relative"]*len(fee_df) + ["total"]
    fig3 = go.Figure(go.Waterfall(
        name = "Dòng tiền",
        orientation = "v",
        measure = measure,
        x = labels,
        y = values,
        text = [f"{v:,.0f}" for v in values],
        connector = {"line": {"color": "orange"}}
    ))
    fig3.update_layout(title="Dòng tiền sau khi trừ các khoản chi phí", height=450)
    st.plotly_chart(fig3, use_container_width=True)

    # Xuất file Excel tổng hợp
    out_xlsx = "tong_hop_bao_cao_pdf.xlsx"
    df_summary.to_excel(out_xlsx, index=False)
    with open(out_xlsx, "rb") as f:
        st.download_button("Tải file tổng hợp Excel", f, out_xlsx, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Vui lòng upload file PDF báo cáo.")

st.caption("Made by ChatGPT – muốn tối ưu thêm, hãy nhắn trực tiếp nhé!")
