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
  — if the PDF lists **more than one rate period**, a separate updated cheatsheet is
  produced **automatically for each period**
""")

st.markdown("#### ⚠️ Important Notes")
st.markdown("""
- If a PDF has **new or removed** POL/POD lanes, remember to add/delete the
  corresponding row in the cheatsheet, and add the matching pair to the **Mapping**
  tab with the **exact** POL/POD spelling as it appears in the PDF extraction
- US inland rows are matched by **Location + Routing via** (e.g. "ATLANTA, GA" + "SAV")
  — if a PDF adds/removes a rail ramp location, add/delete the matching row in the
  **US inland** tab first. If the Location matches but the Routing via doesn't, the
  row is **not** updated and is flagged for you to check instead
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

                file_name = f"COSCO Italy to USA Eff {date.today().strftime('%m-%d-%Y')}.xlsx"

                items = []
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

                    # Read bytes into memory now — the temp file gets deleted
                    # below, but a rerun (e.g. from clicking a download button)
                    # must not lose this data.
                    with open(out_path, "rb") as f:
                        file_bytes = f.read()

                    items.append({"period": period, "result": result, "file_bytes": file_bytes})

                # Persist results in session_state so clicking a download button
                # (which triggers a rerun) doesn't wipe the other results/buttons.
                st.session_state["cosco_results"] = {
                    "rate_ref":  rate_ref,
                    "periods":   periods,
                    "file_name": file_name,
                    "items":     items,
                }

            except ValueError as e:
                st.error(f"❌ Error: {e}")
                st.session_state.pop("cosco_results", None)
            except Exception as e:
                st.error(f"❌ Unexpected error: {e}")
                st.session_state.pop("cosco_results", None)
                raise
            finally:
                for path in [pdf_path, excel_path, *tmp_outputs]:
                    if path:
                        try:
                            os.unlink(path)
                        except Exception:
                            pass
    else:
        st.warning("Please upload both PDF and Cheatsheet.")

# Render persisted results — this runs on every rerun (including the one
# triggered by clicking a download button), so results stay on screen.
if "cosco_results" in st.session_state:
    data = st.session_state["cosco_results"]

    if data["rate_ref"]:
        st.info(f"🏷️ Rate Reference: **{data['rate_ref']}**")
    else:
        st.warning("⚠️ Could not find a Rate Reference value in the PDF")

    if data["periods"]:
        st.info(f"📅 Detected {len(data['periods'])} rate periods in the US inland "
                 f"table: **{', '.join(data['periods'])}** — producing one cheatsheet per period")

    for item in data["items"]:
        period = item["period"]
        result = item["result"]
        label  = f" ({period})" if period else ""

        st.success(f"✅{label} Updated {result['oft_updated']} OFT rows, "
                   f"{result['inland_updated']} US inland rows")

        if result["inland_mismatched"]:
            with st.expander(f"⚠️{label} {len(result['inland_mismatched'])} "
                              f"US inland rows skipped — Location matched but "
                              f"Routing via didn't (not written, please check)"):
                for row_num, location, cs_via, pdf_via in result["inland_mismatched"][:20]:
                    st.text(f"  row {row_num}: Location={location}  "
                            f"cheatsheet Routing via={cs_via}  PDF VIA POD={pdf_via}")

        st.download_button(
            f"📥 Download{label}",
            item["file_bytes"],
            file_name=data["file_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{period}",
        )
