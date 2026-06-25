import streamlit as st
from parsers.cosco_parser import parse_cosco_pdf, parse_wharfage
from excel.excel_writer_test import update_excel_rates_test
import tempfile

st.title("Non TP COSCO SBS Rate Updater")
st.markdown("#### 🔧 What this tool does")
st.markdown("""
- ✅ Copies current rates (W, X, Y) into old rate columns (AD, AE, AF)
- 🚢 Updates Ocean Freight rates (H, I, J)
- ⚓ Updates Wharfage rates (Q, R, S)
""")

st.markdown("#### ⚠️ Important Notes")
st.markdown("""
- ETS is **not updated** (exchange rate dependent)
- Please update the **valid date** manually
- Do **NOT** delete or insert columns (this will break the template)
""")

# 👉 線放這裡（你要的位置）
st.markdown("<br>", unsafe_allow_html=True)

pdf_file = st.file_uploader("Upload pdf", type="pdf")
excel_file = st.file_uploader("Upload excel", type="xlsx") 

if st.button("Run"):

    if pdf_file and excel_file:

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(pdf_file.read())
            pdf_path = tmp_pdf.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
            tmp_excel.write(excel_file.read())
            excel_path = tmp_excel.name

        # ✅ parse
        rates = parse_cosco_pdf(pdf_path)
        whf_rates = parse_wharfage(pdf_path)

        # ✅ update
        update_excel_rates_test(excel_path, rates, whf_rates)

        # ✅ download
        with open(excel_path, "rb") as f:
            st.download_button(
                "Download Updated Excel",
                f,
                file_name="updated_cheatsheet.xlsx"
            )

    else:
        st.warning("Please upload both PDF and Excel.")