import streamlit as st
import os
import subprocess
import fitz  # PyMuPDF
import tempfile
import shutil
from pathlib import Path

# --- CONFIGURATION ---
st.set_page_config(page_title="SlideDeck Merger", page_icon="📑")

# --- UI LAYOUT ---
st.title("📑 Smart Slide Merger")
st.markdown("""
**Upload your Unit PPTs here.** We will convert them to PDF, merge them into one file, and compress it.
*Works on Mac, Windows, and Mobile.*
""")

# --- FILE UPLOADER ---
uploaded_files = st.file_uploader(
    "Drag and drop your PPTX files", 
    type=["pptx"], 
    accept_multiple_files=True
)

# --- BACKEND LOGIC ---
if uploaded_files:
    if st.button("Convert & Merge"):
        
        # Create a temporary directory for this specific user session
        with tempfile.TemporaryDirectory() as temp_dir:
            st.info("Processing... This might take a minute.")
            progress_bar = st.progress(0)
            
            # 1. Save uploaded files to temp folder
            pptx_paths = []
            for uploaded_file in uploaded_files:
                file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                pptx_paths.append(file_path)

            # 2. Sort files (Smart Sort logic)
            # We sort based on numbers found in the filename
            def get_number(filename):
                import re
                match = re.search(r'(\d+)', filename)
                return int(match.group(1)) if match else 999
            
            pptx_paths.sort(key=lambda x: get_number(os.path.basename(x)))

            # 3. Processing Loop
            merger = fitz.open()
            total_files = len(pptx_paths)
            
            # CHECK FOR LIBREOFFICE
            # On Linux (Streamlit Cloud), the command is 'libreoffice'
            # On Windows local, it might be the full path.
            # We try 'libreoffice' first (for cloud), then fallback.
            office_cmd = "libreoffice" 
            
            for i, pptx_path in enumerate(pptx_paths):
                filename = os.path.basename(pptx_path)
                pdf_name = os.path.splitext(filename)[0] + ".pdf"
                pdf_path = os.path.join(temp_dir, pdf_name)
                
                # Update progress
                progress_bar.progress((i / total_files))
                st.write(f"Converting: {filename}...")

                # Conversion Command
                # --headless means no UI, --outdir specifies where to save
                cmd = [
                    office_cmd, '--headless', '--convert-to', 'pdf', 
                    '--outdir', temp_dir, pptx_path
                ]
                
                try:
                    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
                except FileNotFoundError:
                    st.error("Server Error: LibreOffice not found. If running locally, check path.")
                    st.stop()
                except Exception as e:
                    st.error(f"Failed to convert {filename}: {e}")

                # Merge
                if os.path.exists(pdf_path):
                    with fitz.open(pdf_path) as pdf_doc:
                        merger.insert_pdf(pdf_doc)
            
            # 4. Save Final PDF
            output_path = os.path.join(temp_dir, "Merged_Unit.pdf")
            merger.save(output_path, garbage=4, deflate=True)
            merger.close()
            
            progress_bar.progress(100)
            st.success("Conversion Complete!")
            
            # 5. Download Button
            with open(output_path, "rb") as f:
                st.download_button(
                    label="Download Merged PDF",
                    data=f,
                    file_name="Unit_Merged_Complete.pdf",
                    mime="application/pdf"
                )