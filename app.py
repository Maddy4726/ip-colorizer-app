import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill
import re
import tempfile

st.set_page_config(page_title="Excel IP Colorizer", layout="centered")

st.title("📊 Excel IP Colorizer")

# --- File Upload ---
excel_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])
ip_file = st.file_uploader("Upload IP List File (.txt)", type=["txt"])

st.write("### Uploaded Files")
if excel_file:
    st.success(f"Excel file: {excel_file.name}")
if ip_file:
    st.success(f"IP file: {ip_file.name}")


# --- Process Function ---
def process_excel(excel_path, ip_path, output_path):
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    red_fill   = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

    ip_pattern = re.compile(
        r'(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)'
        r'(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}'
    )

    # Read IP list
    with open(ip_path, 'r') as f:
        pinging_ips = {line.strip() for line in f if line.strip()}

    workbook = openpyxl.load_workbook(excel_path)

    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    match = ip_pattern.search(cell.value)
                    if match:
                        ip = match.group()
                        if ip in pinging_ips:
                            cell.fill = green_fill
                        else:
                            cell.fill = red_fill

    workbook.save(output_path)


# --- Button ---
if st.button("🚀 Process Files"):

    if excel_file is None or ip_file is None:
        st.error("Please upload both files.")
    else:
        with st.spinner("Processing..."):
            st.info("Processing all sheets and cells...")

            # Temporary files
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_excel:
                temp_excel.write(excel_file.read())
                excel_path = temp_excel.name

            with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as temp_ip:
                temp_ip.write(ip_file.read())
                ip_path = temp_ip.name

            output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name

            # Run processing
            process_excel(excel_path, ip_path, output_path)

        st.success("✅ Processing complete!")

        # Download button
        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Download Processed File",
                data=f,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


