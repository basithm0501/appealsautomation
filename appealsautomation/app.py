import streamlit as st
import tempfile
import os
import shutil
from AppealAutomation import create_appeals_workbook
import glob

st.title("Appeals Capsheet Generator")

st.write("For internal use by the RUSA Allocations Board.")
st.write("Upload your Appeals Template (as an xlsx file) and Appeals Data File (as a csv file), then select the rows that you would like to process into a capsheet.")

uploaded_template = st.file_uploader("Upload Appeals Template (.xlsx)", type=["xlsx"])
uploaded_csv = st.file_uploader("Upload Appeals Data (.csv)", type=["csv"])
col1, col2 = st.columns(2)
with col1:
    start_row = st.number_input("Start Row", min_value=1, step=1, value=4)
with col2:
    end_row = st.number_input("End Row", min_value=1, step=1, value=start_row)

if st.button("Generate Capsheet"):
    if uploaded_template and uploaded_csv:
        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = os.path.join(tmpdir, uploaded_template.name)
            csv_path = os.path.join(tmpdir, uploaded_csv.name)
            # Save uploaded files to disk
            with open(template_path, "wb") as f:
                f.write(uploaded_template.read())
            with open(csv_path, "wb") as f:
                f.write(uploaded_csv.read())

            # Run your backend
            create_appeals_workbook(template_path=template_path, csv_path=csv_path, start_row=start_row, end_row=end_row)

            # Find the generated file (it may be in the current working directory)
            generated_files = glob.glob("CAP_Workbook_*.xlsx")
            if generated_files:
                # Move the file to tmpdir for download
                result_path = os.path.join(tmpdir, os.path.basename(generated_files[0]))
                shutil.move(generated_files[0], result_path)
                st.success("Capsheet generated!")
                with open(result_path, "rb") as f:
                    st.download_button("Download Capsheet", f, file_name=os.path.basename(result_path))
            else:
                st.error("Capsheet not found. Check backend output.")
    else:
        st.warning("Please upload both files.")
