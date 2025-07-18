import tkinter as tk
from tkinter import ttk
from modules.game_time_tab_poomsae import PoomsaeTab
from modules.game_time_tab_kyorugi import KyorugiTab
from common.version import __version__ as app_version
from common.version import __build_date__ as app_date

class GameTimeCalculator(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title(f"경기 시간 계산기 v{app_version} (빌드: {app_date})")
        self.geometry("1400x750")

        self.master = master
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # 스타일 설정
        style = ttk.Style(self)
        style.configure("TNotebook.Tab", padding=[12, 5], font=('Helvetica', 10))
        style.map("TNotebook.Tab", 
                  background=[("selected", "lightgreen")],
                  foreground=[("selected", "black")])

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # Create Poomsae Tab
        self.poomsae_tab_instance = PoomsaeTab(self.notebook, self)
        self.notebook.add(self.poomsae_tab_instance, text="품새")

        # Create Kyorugi Tab
        self.kyorugi_tab_instance = KyorugiTab(self.notebook, self)
        self.notebook.add(self.kyorugi_tab_instance, text="겨루기")

    def on_close(self):
        self.master.deiconify()
        self.destroy()
