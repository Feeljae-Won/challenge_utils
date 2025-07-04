
import tkinter as tk
from tkinter import filedialog

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
                import pdfplumber
                import pandas as pd

                save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
                if not save_path:
                    self.status_label.config(text="Save cancelled.")
                    return

                all_blocks = []
                with pdfplumber.open(self.file_path) as pdf:
                    for i, page in enumerate(pdf.pages):
                        # Extract blocks of text
                        blocks = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False, use_vertical_writing=False, extra_attrs=["fontname", "size"])
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
                    self.status_label.config(text="No text blocks found in the PDF.")
                    return

                df = pd.DataFrame(all_blocks)
                df.to_excel(save_path, index=False)

                self.status_label.config(text=f"Successfully converted to {save_path}")
            except Exception as e:
                self.status_label.config(text=f"Error: {e}")

if __name__ == "__main__":
    app = PdfToExcelConverter()
    app.mainloop()
