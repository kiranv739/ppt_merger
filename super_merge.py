import os
import threading
import re
import time
import customtkinter as ctk
from tkinter import filedialog, messagebox

# --- BACKEND LIBRARIES ---
import comtypes.client
import pythoncom  # Required for COM threading
import fitz  # PyMuPDF for merging and compression

# --- CONFIGURATION ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class SmartMergeApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Setup
        self.title("Smart Study Suite - PPTX to PDF Merger")
        self.geometry("600x450")
        self.resizable(False, False)

        # UI Layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # 1. Header
        self.lbl_title = ctk.CTkLabel(self, text="Course Unit Merger", font=("Roboto", 24, "bold"))
        self.lbl_title.grid(row=0, column=0, pady=(30, 10), sticky="ew")

        # 2. Folder Selection
        self.btn_select = ctk.CTkButton(self, text="Select Unit Folder", command=self.select_folder, height=40)
        self.btn_select.grid(row=1, column=0, pady=10, padx=50, sticky="ew")

        self.lbl_path = ctk.CTkLabel(self, text="No folder selected", text_color="gray")
        self.lbl_path.grid(row=2, column=0, pady=(0, 20))

        # 3. Status & Progress
        self.lbl_status = ctk.CTkLabel(self, text="Ready", font=("Consolas", 14))
        self.lbl_status.grid(row=3, column=0, pady=10)

        self.progress = ctk.CTkProgressBar(self, orientation="horizontal")
        self.progress.set(0)
        self.progress.grid(row=4, column=0, pady=10, padx=50, sticky="ew")

        # 4. Action Button
        self.btn_run = ctk.CTkButton(self, text="CONVERT & MERGE", command=self.start_process, 
                                     fg_color="#2CC985", hover_color="#229A65", state="disabled", height=50)
        self.btn_run.grid(row=5, column=0, pady=30, padx=50, sticky="ew")

        # State variables
        self.selected_folder = None
        self.pptx_files = []

    def select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.selected_folder = path
            self.lbl_path.configure(text=f"...{path[-40:]}") # Show last 40 chars
            
            # Scan for files immediately to show user what we found
            files = [f for f in os.listdir(path) if f.endswith(".pptx") and not f.startswith("~")]
            self.pptx_files = sorted(files, key=self.extract_number)
            
            if self.pptx_files:
                self.lbl_status.configure(text=f"Found {len(self.pptx_files)} slides (Session {self.extract_number(self.pptx_files[0])} - {self.extract_number(self.pptx_files[-1])})")
                self.btn_run.configure(state="normal")
            else:
                self.lbl_status.configure(text="No .pptx files found in this folder!")
                self.btn_run.configure(state="disabled")

    def extract_number(self, filename):
        # Robust regex to find the session number even if there are spaces
        match = re.search(r'Session\s*(\d+)', filename, re.IGNORECASE)
        return int(match.group(1)) if match else 999

    def start_process(self):
        self.btn_run.configure(state="disabled")
        self.btn_select.configure(state="disabled")
        
        # Run heavy task in a separate thread to keep GUI responsive
        threading.Thread(target=self.process_files, daemon=True).start()

    def process_files(self):
        try:
            # IMPORTANT: Initialize COM for this new thread
            pythoncom.CoInitialize()
            
            ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
            ppt_app.Visible = 1 
            
            total_files = len(self.pptx_files)
            merger = fitz.open() # PyMuPDF object
            
            for index, filename in enumerate(self.pptx_files):
                # Update GUI from thread
                self.update_status(f"Converting: {filename}...", (index / total_files))
                
                input_path = os.path.join(self.selected_folder, filename)
                pdf_name = os.path.splitext(filename)[0] + ".pdf"
                pdf_path = os.path.join(self.selected_folder, pdf_name)
                
                # Convert if PDF doesn't exist
                if not os.path.exists(pdf_path):
                    deck = ppt_app.Presentations.Open(input_path)
                    deck.SaveAs(pdf_path, 32) # 32 = PDF
                    deck.Close()
                
                # Add to Merger
                with fitz.open(pdf_path) as pdf_doc:
                    merger.insert_pdf(pdf_doc)

            # Finalize
            self.update_status("Merging and Compressing...", 0.9)
            output_path = os.path.join(self.selected_folder, "Unit_Merged_Complete.pdf")
            
            # Save with compression (deflate=True, garbage=4 removes unused objects)
            merger.save(output_path, garbage=4, deflate=True)
            merger.close()
            ppt_app.Quit()
            
            self.update_status("Done!", 1.0)
            messagebox.showinfo("Success", f"Merged PDF saved at:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.update_status(f"Error: {str(e)}", 0)
        
        finally:
            # Re-enable buttons
            self.btn_run.configure(state="normal")
            self.btn_select.configure(state="normal")

    def update_status(self, message, progress_val):
        # Helper to update GUI safely
        self.lbl_status.configure(text=message)
        self.progress.set(progress_val)

if __name__ == "__main__":
    app = SmartMergeApp()
    app.mainloop()