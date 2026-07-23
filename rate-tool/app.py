import streamlit as st
import tempfile
import os
from datetime import date
from coscopdf_extract import (
    extract, split_pol, extract_rate_ref,
    extract_us_inland, list_us_inland_periods,
)
from excel_writer import update_cheatsheet

st.title("COSCO Italy to USA rate updater")

st.markdown("#### 🔧 What this tool does")
st.markdown("""
- 📄 Extracts **Ocean Freight rates** from COSCO PDF (Direct Ports + Outports)
- ✍️ Updates **OFT 20' / 40' / 40HC** in the cheatsheet
- 🏷️ Updates the **Rate Reference** (e.g. "TLI GL JULY") in cheatsheet cell **OFT!B1**
- 🚂 Updates **US inland 20DV / 40DV/40HQ** rates from the Rail Ramp table
  — if the PDF lists **more than one rate period** (e.g. APRIL / MAY side by side),
  a separate updated cheatsheet is produced **automatically for each period**
""")

st.markdown("#### ⚠️ Important Notes")
st.markdown("""
- If a PDF has **new or removed** POL/POD lanes, remember to add/delete the
  corresponding row in the cheatsheet, and add the matching pair to the **Mapping**
  tab with the **exact** POL/POD spelling as it appears in the PDF extraction
- US inland rows are matched by **Location** text (e.g. "ATLANTA, GA") — if a PDF
  adds/removes a rail ramp location, add/delete the matching row in the
  **US inland** tab first
""")

st.markdown("<br>", unsafe_allow_html=True)

pdf_file   = st.file_uploader("Upload COSCO PDF", type="pdf")
excel_file = st.file_uploader("Upload Cheatsheet (xlsx, must contain Mapping tab)", type="xlsx")

if st.button("Run"):
    if pdf_file and excel_file:
        with st.spinner("Processing..."):
            pdf_path = excel_path = None
            tmp_outputs = []
            try:
                # Save uploads to temp files
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(pdf_file.read())
                    pdf_path = tmp_pdf.name

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
                    tmp_excel.write(excel_file.read())
                    excel_path = tmp_excel.name

                # Extract the parts that don't depend on rate period
                direct, outports = extract(pdf_path)
                all_rates = split_pol(direct + outports)
                rate_ref  = extract_rate_ref(pdf_path)

                # US inland: detect whether the PDF lists multiple rate periods
                # side by side (e.g. APRIL / MAY). If so, run once per period.
                periods = list_us_inland_periods(pdf_path)
                periods_to_run = periods if periods else [None]

                if rate_ref:
                    st.info(f"🏷️ Rate Reference: **{rate_ref}**")
                else:
                    st.warning("⚠️ Could not find a Rate Reference value in the PDF")

                if periods:
                    st.info(f"📅 Detected {len(periods)} rate periods in the US inland "
                             f"table: **{', '.join(periods)}** — producing one cheatsheet per period")

                today_str = date.today().strftime('%m-%d-%Y')

                for period in periods_to_run:
                    ramp_rows = extract_us_inland(pdf_path, period=period)

                    suffix = f"_{period}" if period else ""
                    out_path = excel_path.replace(".xlsx", f"{suffix}_updated.xlsx")
                    tmp_outputs.append(out_path)

                    result = update_cheatsheet(
                        excel_path, all_rates,
                        rate_ref=rate_ref,
                        ramp_rows=ramp_rows,
                        output_path=out_path,
                    )

                    label = f" ({period})" if period else ""
                    st.success(f"✅{label} Updated {result['oft_updated']} OFT rows, "
                               f"{result['inland_updated']} US inland rows")

                    if result["oft_skipped"]:
                        with st.expander(f"ℹ️{label} {len(result['oft_skipped'])} OFT rows "
                                          f"with no matching rate in PDF (normal for rows "
                                          f"not in the Mapping tab)"):
                            for row_num, por, pod in result["oft_skipped"][:20]:
                                st.text(f"  row {row_num}: POR={por}  POD={pod}")

                    if result["inland_skipped"]:
                        with st.expander(f"ℹ️{label} {len(result['inland_skipped'])} "
                                          f"US inland rows with no matching PDF data"):
                            for row_num, location in result["inland_skipped"][:20]:
                                st.text(f"  row {row_num}: Location={location}")

                    file_name = (f"COSCO Italy to USA Eff {today_str}"
                                 f"{' ' + period if period else ''}.xlsx")
                    with open(out_path, "rb") as f:
                        st.download_button(
                            f"📥 Download{label}",
                            f,
                            file_name=file_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=out_path,
                        )

            except ValueError as e:
                st.error(f"❌ Error: {e}")
            except Exception as e:
                st.error(f"❌ Unexpected error: {e}")
                raise
            finally:
                # Clean up temp files
                for path in [pdf_path, excel_path, *tmp_outputs]:
                    if path:
                        try:
                            os.unlink(path)
                        except Exception:
                            pass
    else:
        st.warning("Please upload both PDF and Cheatsheet.")
