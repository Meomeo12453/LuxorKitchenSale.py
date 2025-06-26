import streamlit as st
import pandas as pd
import pdfplumber
import re
import matplotlib.pyplot as plt
import plotly.graph_objects as go

st.set_page_config(page_title="Dashboard PDF", layout="wide")

st.title("Báo cáo PDF chi phí so với tổng tiền sản phẩm")

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
    if from_date and to_date:
        st.success(f"**Báo cáo từ ngày:** {from_date[:4]}-{from_date[4:6]}-{from_date[6:]}  &nbsp;&nbsp; **đến ngày:** {to_date[:4]}-{to_date[4:6]}-{to_date[6:]}")
    else:
        st.warning("Không nhận diện được thời gian báo cáo trong file PDF.")

    df_summary = parse_pdf_summary(all_text)
    # Loại bỏ các dòng "Báo cáo từ", "đến" khỏi bảng
    df_summary = df_summary[~df_summary["Hạng mục"].str.contains("Báo cáo từ|đến", case=False, na=False)].reset_index(drop=True)

    # Lọc chỉ các mục chi phí: tên có "phí" (không phân biệt hoa/thường)
    fee_df = df_summary[df_summary["Hạng mục"].str.contains("phí", case=False, na=False)].copy()

    # Tìm tổng tiền sản phẩm
    revenue_row = df_summary[df_summary["Hạng mục"].str.contains("Tổng tiền sản phẩm", case=False, na=False)]
    if not revenue_row.empty:
        revenue = float(revenue_row["Giá trị (VND)"].iloc[0])
        fee_df["% trên doanh thu"] = fee_df["Giá trị (VND)"].apply(lambda x: abs(x)/revenue*100 if revenue else 0)
    else:
        st.warning("Không tìm thấy dòng 'Tổng tiền sản phẩm' trong bảng dữ liệu PDF.")
        revenue = None
        fee_df["% trên doanh thu"] = 0.0

    st.subheader("Bảng các loại chi phí & phần trăm so với tổng tiền sản phẩm")
    st.dataframe(fee_df)

    # 1. Bar chart: Chi phí vs % doanh thu
    st.subheader("Biểu đồ cột: Giá trị và phần trăm chi phí trên tổng tiền sản phẩm")
    fig, ax = plt.subplots(figsize=(10,5))
    bars = ax.bar(fee_df["Hạng mục"], abs(fee_df["Giá trị (VND)"]), color='orange')
    for bar, pct in zip(bars, abs(fee_df["% trên doanh thu"])):
        ax.annotate(f"{pct:.1f}%", xy=(bar.get_x() + bar.get_width()/2, bar.get_height()),
                    xytext=(0, 5), textcoords="offset points", ha="center", va="bottom", fontsize=10, color="black", fontweight='bold')
    ax.set_ylabel("Giá trị (VND)")
    ax.set_xlabel("Loại chi phí")
    ax.set_title("Giá trị & % từng loại chi phí so với tổng tiền sản phẩm")
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
    labels = ["Tổng tiền sản phẩm"] + fee_df["Hạng mục"].tolist() + ["Còn lại sau chi phí"]
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

    # Xuất file Excel tổng hợp chỉ gồm các chi phí
    out_xlsx = "tong_hop_chi_phi_vs_doanh_thu.xlsx"
    fee_df.to_excel(out_xlsx, index=False)
    with open(out_xlsx, "rb") as f:
        st.download_button("Tải file chi phí so với doanh thu", f, out_xlsx, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Vui lòng upload file PDF báo cáo.")

st.caption("Made by ChatGPT – tối ưu code theo thực tế, liên hệ khi cần mở rộng thêm các loại chi phí.")
