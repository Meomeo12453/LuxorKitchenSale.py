import streamlit as st
import pandas as pd
import pdfplumber
import re

st.set_page_config(page_title="Dashboard Báo cáo PDF", layout="wide")

st.title("Báo cáo PDF doanh thu & chi phí")
uploaded_pdf = st.file_uploader("Tải lên file PDF báo cáo", type=["pdf"])

def parse_pdf_summary(pdf_text):
    # Tìm dòng tổng kết bằng biểu thức chính quy (regex)
    pattern = r"(?P<item>[^\d\n]+)\s+(?P<amount>[-₫\d,.]+)"
    summary = re.findall(pattern, pdf_text)
    rows = []
    for item, amount in summary:
        # Chỉ lấy các dòng có ý nghĩa
        item = item.strip().replace("\n", " ")
        amount = amount.strip().replace("₫", "").replace(",", "").replace(".", "")
        try:
            # Dấu âm là chi phí
            value = float(amount.replace("-", "")) * (-1 if "-" in amount else 1)
        except:
            continue
        # Loại bỏ các dòng không phải số
        if len(item) < 2 or not re.search(r'\d', amount):
            continue
        rows.append((item, value))
    return pd.DataFrame(rows, columns=["Hạng mục", "Giá trị (VND)"])

if uploaded_pdf:
    with pdfplumber.open(uploaded_pdf) as pdf:
        all_text = ""
        for page in pdf.pages:
            all_text += page.extract_text() + "\n"
    
    st.subheader("Tóm tắt file PDF:")
    st.text_area("Nội dung PDF", all_text[:3000] + " ...", height=350)
    
    # Lấy bảng tổng kết từ PDF
    df_summary = parse_pdf_summary(all_text)
    st.subheader("Tổng hợp chi phí & doanh thu")
    st.dataframe(df_summary)
    
    # Tính phần trăm chi phí so với doanh thu (nếu có)
    revenue_row = df_summary[df_summary["Hạng mục"].str.contains("Tổng tiền sản phẩm", case=False)]
    if not revenue_row.empty:
        revenue = float(revenue_row["Giá trị (VND)"].iloc[0])
        df_summary["% trên doanh thu"] = df_summary["Giá trị (VND)"].apply(lambda x: f"{x/revenue*100:.2f}%" if revenue and x < revenue else "")
    else:
        df_summary["% trên doanh thu"] = ""
    
    # Hiển thị chart bar
    st.subheader("Biểu đồ phần trăm chi phí trên tổng doanh thu")
    showdf = df_summary[(df_summary["Giá trị (VND)"]<0) & (df_summary["% trên doanh thu"]!="")]
    if not showdf.empty:
        st.bar_chart(showdf.set_index("Hạng mục")["% trên doanh thu"].str.replace('%','').astype(float))
    
    # Xuất file Excel tổng hợp
    out_xlsx = "tong_hop_bao_cao_pdf.xlsx"
    df_summary.to_excel(out_xlsx, index=False)
    with open(out_xlsx, "rb") as f:
        st.download_button("Tải file tổng hợp Excel", f, out_xlsx, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Vui lòng upload file PDF báo cáo.")

st.caption("Made by ChatGPT. Để tối ưu cho báo cáo thực tế hoặc dashboard cao cấp hơn, hãy liên hệ!")
