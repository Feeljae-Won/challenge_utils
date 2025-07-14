import tkinter as tk
from tkinter import ttk
from modules.game_time_tab_poomsae import PoomsaeTab
from modules.game_time_tab_kyorugi import KyorugiTab

class GameTimeCalculator(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("경기 시간 계산기")
        self.geometry("1400x750")

        self.master = master
        self.protocol("WM_DELETE_WINDOW", self.on_close)

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
