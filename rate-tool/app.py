import streamlit as st
import tempfile
import os
from datetime import date
from coscopdf_extract import extract
from coscopdf_extract import split_pol
from excel_writer import update_oft_rates

st.title("COSCO Italy to USA rate updater")

st.markdown("#### 🔧 What this tool does")
st.markdown("""
- 📄 Extracts **Ocean Freight rates** from COSCO PDF (Direct Ports + Outports)
- ✍️ Updates **OFT 20' / 40' / 40HC** in the cheatsheet
""")

st.markdown("#### ⚠️ Important Notes")
st.markdown("""
- If a PDF has **new or removed** POL/POD lanes, remember to add/delete the
  corresponding row in the cheatsheet, and add the matching pair to the **Mapping**
  tab with the **exact** POL/POD spelling as it appears in the PDF extraction
""")

st.markdown("<br>", unsafe_allow_html=True)

pdf_file   = st.file_uploader("Upload COSCO PDF", type="pdf")
excel_file = st.file_uploader("Upload Cheatsheet (xlsx, must contain Mapping tab)", type="xlsx")

if st.button("Run"):
    if pdf_file and excel_file:
        with st.spinner("Processing..."):
            try:
                # Save uploads to temp files
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(pdf_file.read())
                    pdf_path = tmp_pdf.name

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
                    tmp_excel.write(excel_file.read())
                    excel_path = tmp_excel.name

                output_path = excel_path.replace(".xlsx", "_updated.xlsx")

                # Extract rates from PDF
                direct, outports = extract(pdf_path)
                combined  = direct + outports
                all_rates = split_pol(combined)   # split POL by '/'

                # Write into cheatsheet
                updated_count, skipped = update_oft_rates(
                    excel_path, all_rates, output_path=output_path
                )

                # Result summary
                st.success(f"✅ Updated {updated_count} rows successfully")

                if skipped:
                    # Only show rows that are actually in the Mapping tab
                    # (skipped = rows in sheet with no matching rate, which is expected for non-mapped rows)
                    st.info(f"ℹ️ {len(skipped)} rows in the cheatsheet had no matching rate in PDF "
                            f"(this is normal for rows not in the Mapping tab)")
                    with st.expander("See skipped rows"):
                        for row_num, por, pod in skipped[:20]:
                            st.text(f"  row {row_num}: POR={por}  POD={pod}")

                # Download button
                with open(output_path, "rb") as f:
                    st.download_button(
                        "📥 Download Updated Cheatsheet",
                        f,
                        file_name=f"COSCO Italy to USA Eff {date.today().strftime('%m-%d-%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except ValueError as e:
                st.error(f"❌ Error: {e}")
            except Exception as e:
                st.error(f"❌ Unexpected error: {e}")
                raise
            finally:
                # Clean up temp files
                for path in [pdf_path, excel_path]:
                    try:
                        os.unlink(path)
                    except Exception:
                        pass
    else:
        st.warning("Please upload both PDF and Cheatsheet.")
