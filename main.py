
import tkinter as tk
from tkinter import filedialog
import os
import sys
import pdfplumber
import pandas as pd
from PIL import Image
import pytesseract

# Set the path to the tesseract executable dynamically for PyInstaller bundled app
if getattr(sys, 'frozen', False):
    # Running in a PyInstaller bundle
    application_path = sys._MEIPASS
else:
    # Running in a normal Python environment
    application_path = os.path.dirname(os.path.abspath(__file__))

# Assuming tesseract.exe is bundled in the root of the extracted directory
pytesseract.pytesseract.tesseract_cmd = os.path.join(application_path, 'tesseract.exe')

class PdfToExcelConverter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF to Excel Converter")
        self.geometry("400x200")

        self.file_path = None

        self.select_button = tk.Button(self, text="Select PDF", command=self.select_pdf)
        self.select_button.pack(pady=20)

        self.convert_button = tk.Button(self, text="Convert to Excel", command=self.convert_to_excel, state=tk.DISABLED)
        self.convert_button.pack(pady=10)

        self.status_label = tk.Label(self, text="")
        self.status_label.pack()

    def select_pdf(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.file_path:
            self.status_label.config(text=f"Selected file: {self.file_path}")
            self.convert_button.config(state=tk.NORMAL)

    def convert_to_excel(self):
        if self.file_path:
            try:
                # Generate default save name from the selected PDF file
                base_name = os.path.basename(self.file_path)
                file_name_without_ext = os.path.splitext(base_name)[0]
                default_save_name = f"{file_name_without_ext}_엑셀 변환"

                save_path = filedialog.asksaveasfilename(
                    initialfile=default_save_name,
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")]
                )

                if not save_path:
                    self.status_label.config(text="Save cancelled.")
                    return

                all_blocks = []
                with pdfplumber.open(self.file_path) as pdf:
                    for i, page in enumerate(pdf.pages):
                        # Try to extract text blocks using pdfplumber first
                        blocks = page.extract_words()

                        if not blocks: # If no text blocks found, try OCR
                            self.status_label.config(text=f"Page {i+1}: No text blocks found. Attempting OCR...")
                            try:
                                # Convert PDF page to image using pdfplumber's .to_image()
                                # This requires Ghostscript to be installed and in PATH.
                                img = page.to_image(resolution=300).original
                                ocr_text = pytesseract.image_to_string(img, lang='kor+eng') # 'kor' for Korean, 'eng' for English
                                if ocr_text.strip():
                                    # For simplicity, treat the whole OCR'd text as one block
                                    all_blocks.append({
                                        'page': i + 1,
                                        'text': ocr_text.strip(),
                                        'x0': 0, 'top': 0, 'x1': 0, 'bottom': 0, # Placeholder coordinates
                                        'fontname': 'OCR',
                                        'size': 0
                                    })
                                else:
                                    self.status_label.config(text=f"Page {i+1}: OCR found no text.")
                            except pytesseract.TesseractNotFoundError:
                                self.status_label.config(text="Error: Tesseract-OCR is not installed or not in PATH. Please install it from https://tesseract-ocr.github.io/tessdoc/Installation.html and ensure it's in your system PATH.")
                                return
                            except Exception as ocr_e:
                                self.status_label.config(text=f"Page {i+1}: OCR Error: {ocr_e}. Ensure Ghostscript is installed and in PATH for image conversion.")
                                continue # Skip to next page if OCR fails
                        else:
                            for block in blocks:
                                all_blocks.append({
                                    'page': i + 1,
                                    'text': block['text'],
                                    'x0': block['x0'],
                                    'top': block['top'],
                                    'x1': block['x1'],
                                    'bottom': block['bottom'],
                                    'fontname': block.get('fontname', ''),
                                    'size': block.get('size', 0)
                                })

                if not all_blocks:
                    self.status_label.config(text="No text blocks found in the PDF, even with OCR.")
                    return

                df = pd.DataFrame(all_blocks)
                df.to_excel(save_path, index=False)

                self.status_label.config(text=f"Successfully converted to {save_path}")
            except Exception as e:
                self.status_label.config(text=f"Error: {e}")

if __name__ == "__main__":
    app = PdfToExcelConverter()
    app.mainloop()
