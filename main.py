import tkinter as tk
from tkinter import font
import datetime
from version import __version__ as app_version
from version import __build_date__ as app_date

class PasswordWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("비밀번호 입력")
        self.geometry("300x150")
        self.resizable(False, False)

        self.password_label = tk.Label(self, text="비밀번호를 입력하세요:")
        self.password_label.pack(pady=10)

        self.password_entry = tk.Entry(self, show="*")
        self.password_entry.pack(pady=5)
        self.password_entry.focus_set() # 이 줄을 추가합니다.
        self.password_entry.bind("<Return>", self.check_password) # Enter 키 바인딩

        self.login_button = tk.Button(self, text="확인", command=self.check_password)
        self.login_button.pack(pady=10)

        # 창을 화면 중앙에 배치
        self.update_idletasks()
        x = self.winfo_screenwidth() // 2 - self.winfo_width() // 2
        y = self.winfo_screenheight() // 2 - self.winfo_height() // 2
        self.geometry(f"300x150+{x}+{y}")

    def check_password(self, event=None):
        entered_password = self.password_entry.get()
        if entered_password == "015394":
            self.destroy() # 비밀번호 창 닫기
            app = MainApp() # 메인 앱 실행
            app.mainloop()
        else:
            tk.messagebox.showerror("오류", "잘못된 비밀번호입니다.")
            self.password_entry.delete(0, tk.END) # 입력 필드 초기화


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.version = app_version # 변경
        self.title(f"필재의 유틸리티 모음 v{app_version} (빌드: {app_date}")
        self.geometry("400x500")

        # 메인 타이틀
        title_font = font.Font(family="Helvetica", size=24, weight="bold")
        title_label = tk.Label(self, text="유틸 목록", font=title_font, pady=20)
        title_label.pack()

        # 버전 정보 레이블
        version_font = font.Font(family="Helvetica", size=10, slant="italic")
        version_label = tk.Label(self, text=f"Version {self.version} (Build: {app_date})", font=version_font, fg="gray")
        version_label.pack()

        # 모듈 이동 버튼 프레임
        button_frame = tk.Frame(self)
        button_frame.pack(pady=10)

        # 모듈 버튼 생성
        modules = {
            "경기번호 계산기": self.open_game_number_calculator,
            "품새 소청 계산기": self.open_poomsae_sochung_calculator,
            "준비중": lambda: self.on_module_button_click("PDF 변환")
        }
        for module_name, command in modules.items():
            self.create_module_button(button_frame, module_name, command)

        # Footer
        footer_frame = tk.Frame(self)
        footer_frame.pack(side=tk.BOTTOM, pady=5)

        notes_font = font.Font(family="Helvetica", size=9, underline=True)
        notes_button = tk.Button(footer_frame, text="Notes", font=notes_font, fg="blue", relief="flat", cursor="hand2", command=self.open_release_notes)
        notes_button.pack(side=tk.LEFT, padx=5)

        footer_font = font.Font(family="Helvetica", size=9)
        footer_label = tk.Label(footer_frame, text="Copyright (c) FEELJAE-WON. All rights reserved.", font=footer_font, fg="gray")
        footer_label.pack(side=tk.LEFT, padx=5)

    def create_module_button(self, parent, module_name, command):
        button_font = font.Font(family="Helvetica", size=12)
        button = tk.Button(parent, text=module_name, font=button_font, width=30, height=2,
                           command=command)
        button.pack(pady=8)

    def on_module_button_click(self, module_name):
        print(f"'{module_name}' 모듈로 이동합니다.")

    def open_release_notes(self):
        release_notes_window = tk.Toplevel(self)
        release_notes_window.title("릴리스 노트")
        release_notes_window.geometry("600x400")

        text_area = tk.Text(release_notes_window, wrap="word", padx=10, pady=10)
        text_area.pack(expand=True, fill="both")

        try:
            with open("RELEASE_NOTES.md", "r", encoding="utf-8") as f:
                release_notes_content = f.read()
            text_area.insert("1.0", release_notes_content)
        except FileNotFoundError:
            text_area.insert("1.0", "릴리스 노트 파일을 찾을 수 없습니다.")
        
        text_area.config(state="disabled")

    def open_game_number_calculator(self):
        from modules.game_number_calculator import GameNumberCalculator
        self.withdraw() # 메인 창 숨기기
        calculator_window = GameNumberCalculator(self)
        calculator_window.protocol("WM_DELETE_WINDOW", lambda: self.on_calculator_close(calculator_window))

    def on_calculator_close(self, window):
        window.destroy()
        self.deiconify() # 메인 창 다시 보이기

    def open_poomsae_sochung_calculator(self):
        from modules.poomsae_sochung_calculator import PoomsaeSochungCalculator
        self.withdraw() # 메인 창 숨기기
        calculator_window = PoomsaeSochungCalculator(self)
        calculator_window.protocol("WM_DELETE_WINDOW", lambda: self.on_calculator_close(calculator_window))

if __name__ == "__main__":
    password_app = PasswordWindow()
    password_app.mainloop()