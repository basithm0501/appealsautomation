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

# How to Use Section
st.markdown("---")
st.header("📋 How to Use This Tool - Step by Step Guide")

st.markdown("""
This tool helps you generate appeals capsheets automatically. Follow these detailed steps carefully:

### Step 1: Get the Appeals Template File
1. **Locate the most recent and updated appeals capsheet template**
   - This should be an Excel file with the extension `.xlsx`
   - Make sure you have the latest version of the template
   - The file should contain all the necessary formatting and columns for appeals processing
   - If you don't have the template, contact your supervisor or the person who manages the appeals process

### Step 2: Download and Prepare Your Appeals Data
1. **Download the appeals form data from GetInvolved**
   - Log into the GetInvolved platform
   - Navigate to the appeals section or form submissions
   - Download the data file (this will typically be in CSV or Excel format)

2. **Open the downloaded file in a spreadsheet application**
   - You can use:
     - **Numbers** (Mac users)
     - **Microsoft Excel** (Windows/Mac users)
     - **Google Sheets** (web browser - any device)
   - Double-click the file to open it, or use File → Open in your chosen application

3. **Clean up the data**
   - **Delete all the rows you don't need**
   - Keep only the appeals you want to process
   - Make sure to keep the header row (the first row with column names)
   - Remove any empty rows or irrelevant entries

4. **Export as a CSV file**
   - Go to **File** → **Export As** (or **Save As** in some applications)
   - Choose **CSV** format (Comma Separated Values)
   - **Important**: Give it a clear name and save it somewhere you can easily find it
   - Remember the location where you saved this file

### Step 3: Identify Row Numbers
1. **Keep note of your data rows**
   - **First row with data**: Look at your cleaned CSV file and count which row number contains your first appeal (usually row 2, since row 1 is typically headers)
   - **Last row with data**: Count down to find the row number of your last appeal
   - **Example**: If you have headers in row 1 and appeals in rows 2 through 25, then:
     - Start Row = 2
     - End Row = 25

### Step 4: Use This Website Tool
1. **Upload your files**
   - Click "Browse files" under **"Upload Appeals Template (.xlsx)"**
   - Select and upload your appeals template Excel file
   - Click "Browse files" under **"Upload Appeals Data (.csv)"**
   - Select and upload your cleaned CSV file

2. **Enter the row numbers**
   - In the **"Start Row"** field: Enter the row number where your first appeal data begins
   - In the **"End Row"** field: Enter the row number where your last appeal data ends

3. **Generate your capsheet**
   - Click the **"Generate Capsheet"** button
   - **Wait patiently** - the processing may take a few moments depending on the size of your data
   - Do not refresh the page or close the browser while it's processing

4. **Download your result**
   - Once processing is complete, you'll see a "Download Capsheet" button
   - Click it to download your generated appeals capsheet
   - The file will be saved to your computer's Downloads folder (or wherever your browser saves files)

### ⚠️ Important Tips:
- **File formats matter**: Template must be `.xlsx`, data must be `.csv`
- **Double-check your row numbers**: Incorrect row numbers will cause errors or incomplete processing
- **Keep your files organized**: Save everything in a folder you can easily find
- **Don't close the browser**: Wait for the "Download Capsheet" button to appear before doing anything else
- **If you get an error**: Check that both files are uploaded correctly and row numbers are valid

### 🆘 Troubleshooting:
- **"Please upload both files" message**: Make sure you've selected both the template (.xlsx) and data (.csv) files
- **"Capsheet not found" error**: Check your row numbers and ensure your CSV file has data in those rows
- **Processing takes too long**: Large files may take several minutes - be patient and don't refresh
- **Download doesn't work**: Try right-clicking the download button and selecting "Save link as"

### 📞 Need Help?
If you're still having trouble, contact me.
""")
