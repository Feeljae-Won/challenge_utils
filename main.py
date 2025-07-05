import tkinter as tk
from tkinter import font

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("필재의 유틸리티 모음")
        self.geometry("400x500")

        # 메인 타이틀
        title_font = font.Font(family="Helvetica", size=24, weight="bold")
        title_label = tk.Label(self, text="유틸 목록", font=title_font, pady=20)
        title_label.pack()

        # 모듈 이동 버튼 프레임
        button_frame = tk.Frame(self)
        button_frame.pack(pady=10)

        # 모듈 버튼 생성
        modules = {
            "경기번호 계산기": self.open_game_number_calculator,
            "준비중": lambda: self.on_module_button_click("PDF 변환")
        }
        for module_name, command in modules.items():
            self.create_module_button(button_frame, module_name, command)

        # Footer
        footer_font = font.Font(family="Helvetica", size=9)
        footer_label = tk.Label(self, text="Copyright (c) FEELJAE-WON. All rights reserved.", font=footer_font, fg="gray")
        footer_label.pack(side=tk.BOTTOM, pady=5)

    def create_module_button(self, parent, module_name, command):
        button_font = font.Font(family="Helvetica", size=12)
        button = tk.Button(parent, text=module_name, font=button_font, width=30, height=2,
                           command=command)
        button.pack(pady=8)

    def on_module_button_click(self, module_name):
        print(f"'{module_name}' 모듈로 이동합니다.")

    def open_game_number_calculator(self):
        from modules.game_number_calculator import GameNumberCalculator
        self.withdraw() # 메인 창 숨기기
        calculator_window = GameNumberCalculator(self)
        calculator_window.protocol("WM_DELETE_WINDOW", lambda: self.on_calculator_close(calculator_window))

    def on_calculator_close(self, window):
        window.destroy()
        self.deiconify() # 메인 창 다시 보이기

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()