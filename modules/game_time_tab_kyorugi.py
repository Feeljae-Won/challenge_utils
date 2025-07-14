import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import shutil
import os
import openpyxl
from datetime import datetime, timedelta

SETTINGS_FILE = "kyorugi_settings.json"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), '..', 'templates', '겨루기_경기시간_계산기_양식.xlsx')

DEFAULT_SETTINGS = {
    "round_time": "120", # Default round time in seconds
    "break_time": "30",  # Default break time in seconds
    "golden_round_time": "60", # Default golden round time in seconds
    # Add more default settings for Kyorugi here as needed
}

class KyorugiTab(ttk.Frame):
    def __init__(self, notebook, parent_app):
        super().__init__(notebook)
        self.parent_app = parent_app
        self.input_rows = []
        self.create_widgets()
        self.populate_default_rows() # Populate with default rows for Kyorugi

    def create_widgets(self):
        main_paned_window = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_paned_window.pack(fill=tk.BOTH, expand=True)

        left_frame = ttk.Frame(main_paned_window, width=550)
        right_frame = ttk.Frame(main_paned_window, width=350)
        main_paned_window.add(left_frame, weight=11)
        main_paned_window.add(right_frame, weight=7)

        control_frame = tk.Frame(left_frame)
        control_frame.pack(fill='x', padx=10, pady=(10,0))

        add_button = tk.Button(control_frame, text="+", command=self.add_input_row)
        add_button.pack(side="left", padx=(0, 5))
        import_button = tk.Button(control_frame, text="엑셀로 가져오기", command=self.import_from_excel)
        import_button.pack(side="left", padx=5)
        download_button = tk.Button(control_frame, text="엑셀 양식 다운로드", command=self.download_excel_template)
        download_button.pack(side="left", padx=5)
        settings_button = tk.Button(control_frame, text="⚙️ 옵션", command=self.open_kyorugi_settings)
        settings_button.pack(side="right")

        input_grid_frame = tk.Frame(left_frame)
        input_grid_frame.pack(expand=True, fill="both", padx=10, pady=10)

        header_frame = tk.Frame(input_grid_frame)
        header_frame.pack(fill='x', pady=(5, 5))
        self.header_check_var = tk.IntVar()
        header_check = tk.Checkbutton(header_frame, variable=self.header_check_var, command=self.toggle_all_checks)
        header_check.pack(side="left", padx=2, anchor='w')
        tk.Label(header_frame, text="종목", width=15, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="참가부", width=15, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="세부부별", width=15, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="성별", width=8, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="인원수", width=10, anchor='w').pack(side="left", padx=2)

        canvas = tk.Canvas(input_grid_frame)
        scrollbar = tk.Scrollbar(input_grid_frame, orient="vertical", command=canvas.yview)
        self.rows_container = tk.Frame(canvas)
        self.rows_container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.rows_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        results_labelframe = tk.LabelFrame(right_frame, text="결과")
        results_labelframe.pack(expand=True, fill="both", padx=10, pady=10)

        top_input_frame = tk.Frame(results_labelframe)
        top_input_frame.pack(fill='x', padx=5, pady=5)

        time_frame = tk.Frame(top_input_frame)
        time_frame.pack(fill='x')
        tk.Label(time_frame, text="시작 시간:").pack(side="left")
        self.start_time_var = tk.StringVar(value="09:00")
        time_entry = tk.Entry(time_frame, textvariable=self.start_time_var, width=8)
        time_entry.pack(side="left", padx=(5,2))
        now_button = tk.Button(time_frame, text="현재 시간", command=self.set_current_time)
        now_button.pack(side="left")

        court_frame = tk.Frame(top_input_frame)
        court_frame.pack(fill='x', pady=(5,0))
        tk.Label(court_frame, text="코트수:").pack(side="left", padx=(0,5))
        self.court_entry = tk.Entry(court_frame, width=5)
        self.court_entry.insert(0, "4") # Default court count for Kyorugi
        self.court_entry.pack(side="left", padx=(0, 10))

        calc_button = tk.Button(results_labelframe, text="계산하기", bg="red", fg="white", font=("Helvetica", 10, "bold"), command=self.calculate_time)
        calc_button.pack(fill='x', padx=5, pady=10)
        self.bind('<Return>', lambda event=None: calc_button.invoke())

        self.result_text = tk.Text(results_labelframe, height=15, wrap="word", state="disabled", relief="flat")
        self.result_text.pack(expand=True, fill="both", padx=5, pady=5)

    def calculate_time(self):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                settings = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            settings = DEFAULT_SETTINGS

        try:
            start_time = datetime.strptime(self.start_time_var.get(), "%H:%M")
            court_count = int(self.court_entry.get())
            if court_count <= 0:
                raise ValueError("코트 수는 0보다 커야 합니다.")
        except ValueError as e:
            messagebox.showerror("입력 오류", f"시작 시간 또는 코트 수 입력이 잘못되었습니다.\n{e}", parent=self)
            return

        selected_rows = [row for row in self.input_rows if row['check_var'].get() == 1]
        rows_to_process = selected_rows if selected_rows else self.input_rows

        total_kyorugi_seconds_raw = 0
        # sub_totals for Kyorugi can be defined here if needed, e.g., by weight class or gender

        for row in rows_to_process:
            try:
                event_input = row['event'].get()
                division_input = row['division'].get()
                category = row['class'].get()
                gender = row['gender'].get()
                original_headcount = int(row['count'].get() or 0)
                
                # Kyorugi specific calculation logic (placeholder)
                # For now, let's assume each person takes a fixed time for a match
                # This needs to be defined based on user requirements for Kyorugi
                time_per_match = int(settings.get("round_time", 120)) + int(settings.get("break_time", 30)) # Example
                
                row_total_seconds = original_headcount * time_per_match
                total_kyorugi_seconds_raw += row_total_seconds

            except (ValueError, KeyError) as e:
                messagebox.showerror("데이터 오류", f"입력 데이터에 오류가 있습니다. 확인해주세요.\n종목: {event_input}, 참가부: {division_input}, 세부부별: {category}, 성별: {gender}\n오류: {e}", parent=self)
                return

        kyorugi_duration_per_court = total_kyorugi_seconds_raw / court_count if court_count > 0 else 0

        total_duration_seconds = kyorugi_duration_per_court
        end_time = start_time + timedelta(seconds=total_duration_seconds)

        # --- Display Results ---
        def format_time(seconds):
            m, s = divmod(seconds, 60)
            h, m = divmod(m, 60)
            return f"{int(h)}시간 {int(m)}분 {int(s)}초"

        result_str = "===== 코트 적용 소요시간 ====\n"
        result_str += f"총 예상 소요시간: {format_time(total_duration_seconds)}\n"
        result_str += "\n[겨루기] - " + str(court_count) + " 코트 기준\n"
        result_str += f"  총 소요시간: {format_time(kyorugi_duration_per_court)}\n"

        result_str += "===========\n"
        result_str += f"시작 시간: {start_time.strftime('%H:%M')}\n"
        result_str += f"예상 종료 시간: {end_time.strftime('%H:%M')}\n"

        self.result_text.config(state="normal")
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, result_str)
        self.result_text.config(state="disabled")

    def set_current_time(self):
        now = datetime.now()
        current_time = now.strftime("%H:%M")
        self.start_time_var.set(current_time)

    def populate_default_rows(self):
        # Default rows for Kyorugi (can be customized later)
        for i in range(len(self.input_rows) - 1, -1, -1):
            self.input_rows[i]['frame'].destroy()
            self.input_rows.pop(i)

        # Example default rows for Kyorugi
        all_rows_data = []
        for category in ["초등부", "중등부", "고등부", "대학부", "일반부"]:
            all_rows_data.append({"종목": "겨루기", "참가부": category, "세부부별": "", "성별": "남자", "인원수": ""})
            all_rows_data.append({"종목": "겨루기", "참가부": category, "세부부별": "", "성별": "여자", "인원수": ""})

        # Sort the rows (can be customized for Kyorugi)
        division_order = {"초등부": 1, "중등부": 2, "고등부": 3, "대학부": 4, "일반부": 5}
        all_rows_data.sort(key=lambda x: division_order.get(x["참가부"], 99))

        for data in all_rows_data:
            self.add_input_row(data)

        if not self.input_rows:
            self.add_input_row()

    def toggle_all_checks(self):
        is_checked = self.header_check_var.get()
        for row_widgets in self.input_rows:
            row_widgets['check_var'].set(is_checked)

    def add_input_row(self, data=None):
        row_frame = tk.Frame(self.rows_container)
        row_frame.pack(fill='x', pady=2, anchor='w')

        check_var = tk.IntVar()
        check = tk.Checkbutton(row_frame, variable=check_var)
        check.pack(side="left", padx=2)

        event_combo = ttk.Combobox(row_frame, values=["겨루기"], width=15) # Kyorugi specific values
        event_combo.pack(side="left", padx=2)

        division_combo = ttk.Combobox(row_frame, values=["초등부", "중등부", "고등부", "대학부", "일반부"], width=15)
        division_combo.pack(side="left", padx=2)

        category_entry = tk.Entry(row_frame, width=15)
        category_entry.pack(side="left", padx=2)

        gender_combo = ttk.Combobox(row_frame, values=["남자", "여자", "혼성"], width=8)
        gender_combo.pack(side="left", padx=2)

        count_entry = tk.Entry(row_frame, width=10)
        count_entry.pack(side="left", padx=2)
        
        count_entry.bind("<Tab>", self.focus_next_game_count)
        count_entry.bind("<Shift-Tab>", self.focus_prev_game_count)

        delete_button = tk.Button(row_frame, text="-", command=lambda: self.remove_input_row(row_widgets))
        delete_button.pack(side="left", padx=2)

        row_widgets = {
            'frame': row_frame, 'check_var': check_var, 'event': event_combo, 
            'division': division_combo, 'class': category_entry, 'gender': gender_combo, 'count': count_entry
        }
        self.input_rows.append(row_widgets)

        if data:
            event_combo.set(data.get("종목", ""))
            division_combo.set(data.get("참가부", ""))
            category_entry.insert(0, data.get("세부부별", ""))
            gender_combo.set(data.get("성별", ""))
            count_entry.insert(0, str(data.get("인원수", "")))

    def focus_next_game_count(self, event):
        current_entry = event.widget
        for i, row_widgets in enumerate(self.input_rows):
            if row_widgets['count'] == current_entry:
                next_index = (i + 1) % len(self.input_rows)
                self.input_rows[next_index]['count'].focus_set()
                self.input_rows[next_index]['count'].selection_range(0, tk.END) # Select all text
                return "break" # Prevent default tab behavior
        return "continue" # Allow default tab behavior if not found

    def focus_prev_game_count(self, event):
        current_entry = event.widget
        for i, row_widgets in enumerate(self.input_rows):
            if row_widgets['count'] == current_entry:
                prev_index = (i - 1 + len(self.input_rows)) % len(self.input_rows)
                self.input_rows[prev_index]['count'].focus_set()
                self.input_rows[prev_index]['count'].selection_range(0, tk.END) # Select all text
                return "break" # Prevent default tab behavior
        return "continue" # Allow default tab behavior if not found

    def remove_input_row(self, row_widgets):
        if len(self.input_rows) > 1:
            row_widgets['frame'].destroy()
            self.input_rows.remove(row_widgets)
        else:
            messagebox.showwarning("삭제 불가", "마지막 행은 삭제할 수 없습니다.", parent=self)

    def import_from_excel(self):
        file_path = filedialog.askopenfilename(
            parent=self,
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not file_path:
            return

        try:
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active

            for i in range(len(self.input_rows) - 1, -1, -1):
                self.input_rows[i]['frame'].destroy()
                self.input_rows.pop(i)

            for row in sheet.iter_rows(min_row=2, values_only=True):
                if all(cell is None or str(cell).strip() == "" for cell in row):
                    continue
                data = {"종목": row[0], "참가부": row[1], "세부부별": row[2], "성별": row[3], "인원수": row[4]}
                self.add_input_row(data)
            
            if not self.input_rows: 
                self.add_input_row()

        except Exception as e:
            messagebox.showerror("가져오기 실패", f"엑셀 파일을 읽는 중 오류가 발생했습니다:\n{e}", parent=self)

    def download_excel_template(self):
        save_path = filedialog.asksaveasfilename(
            parent=self,
            title="엑셀 양식 저장",
            initialfile="겨루기_경기시간_계산_양식.xlsx", # Use Kyorugi specific template
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.* ")]
        )
        if save_path:
            try:
                shutil.copyfile(TEMPLATE_PATH, save_path)
            except Exception as e:
                messagebox.showerror("저장 실패", f"양식을 저장하는 중 오류가 발생했습니다:\n{e}", parent=self)

    def open_kyorugi_settings(self):
        settings_window = tk.Toplevel(self)
        settings_window.title("겨루기 옵션")
        settings_window.geometry("400x300")
        settings_window.transient(self)
        settings_window.grab_set()

        self.entries = {}

        main_frame = tk.Frame(settings_window)
        main_frame.pack(padx=15, pady=15, fill="both", expand=True)

        # --- Round Time ---
        round_time_frame = tk.Frame(main_frame)
        round_time_frame.pack(fill="x", pady=(0, 10))
        tk.Label(round_time_frame, text="라운드 시간 (초):").pack(side="left", padx=(0, 5))
        self.entries['round_time'] = tk.Entry(round_time_frame, width=10)
        self.entries['round_time'].pack(side="left")

        # --- Break Time ---
        break_time_frame = tk.Frame(main_frame)
        break_time_frame.pack(fill="x", pady=(0, 10))
        tk.Label(break_time_frame, text="휴식 시간 (초):").pack(side="left", padx=(0, 5))
        self.entries['break_time'] = tk.Entry(break_time_frame, width=10)
        self.entries['break_time'].pack(side="left")

        # --- Golden Round Time ---
        golden_round_time_frame = tk.Frame(main_frame)
        golden_round_time_frame.pack(fill="x", pady=(0, 10))
        tk.Label(golden_round_time_frame, text="골든 라운드 시간 (초):").pack(side="left", padx=(0, 5))
        self.entries['golden_round_time'] = tk.Entry(golden_round_time_frame, width=10)
        self.entries['golden_round_time'].pack(side="left")

        # --- Buttons ---
        button_frame = tk.Frame(settings_window)
        button_frame.pack(fill='x', side='bottom', pady=10, padx=15)
        tk.Button(button_frame, text="저장", command=lambda: self.save_settings(settings_window)).pack(side="right", padx=(5,0))
        tk.Button(button_frame, text="기본값", command=lambda: self.restore_defaults(settings_window)).pack(side="right")

        self.load_settings()

    def _get_settings_from_ui(self):
        settings = {
            'round_time': self.entries['round_time'].get(),
            'break_time': self.entries['break_time'].get(),
            'golden_round_time': self.entries['golden_round_time'].get(),
        }
        return settings

    def _save_logic(self, settings_to_save):
        try:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(settings_to_save, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            messagebox.showerror("저장 실패", f"설정을 저장하는 중 오류가 발생했습니다:\n{e}", parent=self)
            return False

    def save_settings(self, window):
        current_settings = self._get_settings_from_ui()
        if self._save_logic(current_settings):
            window.destroy()

    def restore_defaults(self, window):
        if self._save_logic(DEFAULT_SETTINGS):
            window.destroy()

    def load_settings(self):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                settings_to_load = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            settings_to_load = DEFAULT_SETTINGS

        self.entries['round_time'].delete(0, tk.END)
        self.entries['round_time'].insert(0, settings_to_load.get('round_time', DEFAULT_SETTINGS['round_time']))
        self.entries['break_time'].delete(0, tk.END)
        self.entries['break_time'].insert(0, settings_to_load.get('break_time', DEFAULT_SETTINGS['break_time']))
        self.entries['golden_round_time'].delete(0, tk.END)
        self.entries['golden_round_time'].insert(0, settings_to_load.get('golden_round_time', DEFAULT_SETTINGS['golden_round_time']))
