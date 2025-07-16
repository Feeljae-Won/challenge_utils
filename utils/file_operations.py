import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import sys

def download_template_file(template_path, initial_file_name, filetypes):
    if getattr(sys, 'frozen', False):
        # PyInstaller로 번들된 경우
        source_path = os.path.join(sys._MEIPASS, "templates", os.path.basename(template_path))
    else:
        # 개발 환경인 경우
        source_path = template_path

    if not os.path.exists(source_path):
        messagebox.showerror("오류", "양식 파일이 존재하지 않습니다.")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=initial_file_name,
        filetypes=filetypes
    )
    if save_path:
        try:
            shutil.copyfile(source_path, save_path)
            messagebox.showinfo("성공", f"양식이 {save_path}에 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("저장 실패", f"양식을 저장하는 중 오류가 발생했습니다:\n{e}")
