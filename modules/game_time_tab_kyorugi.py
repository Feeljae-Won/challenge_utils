import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import shutil
import os
import openpyxl
from datetime import datetime, timedelta

SETTINGS_FILE = "kyorugi_settings.json"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), '..', 'templates', '겨루기_경기시간_계산기_양식.xlsx')

class KyorugiTab(ttk.Frame):
    def __init__(self, notebook, parent_app):
        super().__init__(notebook)
        self.parent_app = parent_app
        self.input_rows = []
        self.settings_entries = {}
        self.create_widgets()
        self.populate_default_rows()

    def create_widgets(self):
        main_paned_window = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_paned_window.pack(fill=tk.BOTH, expand=True)

        left_frame = ttk.Frame(main_paned_window, width=550)
        right_frame = ttk.Frame(main_paned_window, width=350)
        main_paned_window.add(left_frame, weight=11)
        main_paned_window.add(right_frame, weight=7)

        control_frame = tk.Frame(left_frame)
        control_frame.pack(fill='x', padx=10, pady=(10,0))

        add_button = tk.Button(control_frame, text="+ 개발", command=self.add_input_row)
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
        tk.Label(header_frame, text="참가부", width=20, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="체급", width=20, anchor='w').pack(side="left", padx=2)
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
        self.court_entry.insert(0, "4")
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
            settings = {}

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
        
        for row in rows_to_process:
            try:
                division = row['division'].get()
                weight_class = row['weight_class'].get()
                headcount = int(row['count'].get() or 0)
                
                if not division or headcount == 0:
                    continue

                time_per_match = int(settings.get(division, 450)) # 기본값 450초
                
                num_matches = max(0, headcount - 1)
                
                row_total_seconds = num_matches * time_per_match
                total_kyorugi_seconds_raw += row_total_seconds

            except (ValueError, KeyError) as e:
                messagebox.showerror("데이터 오류", f"입력 데이터에 오류가 있습니다. 확인해주세요.\n참가부: {division}, 체급: {weight_class}\n오류: {e}", parent=self)
                return

        kyorugi_duration_per_court = total_kyorugi_seconds_raw / court_count if court_count > 0 else 0

        total_duration_seconds = kyorugi_duration_per_court
        end_time = start_time + timedelta(seconds=total_duration_seconds)

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
        for i in range(len(self.input_rows) - 1, -1, -1):
            self.input_rows[i]['frame'].destroy()
            self.input_rows.pop(i)
        
        for _ in range(10):
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

        division_entry = tk.Entry(row_frame, width=22)
        division_entry.pack(side="left", padx=2)

        weight_class_entry = tk.Entry(row_frame, width=22)
        weight_class_entry.pack(side="left", padx=2)

        count_entry = tk.Entry(row_frame, width=12)
        count_entry.pack(side="left", padx=2)
        
        delete_button = tk.Button(row_frame, text="-", command=lambda: self.remove_input_row(row_widgets))
        delete_button.pack(side="left", padx=2)

        row_widgets = {
            'frame': row_frame, 'check_var': check_var, 'division': division_entry, 
            'weight_class': weight_class_entry, 'count': count_entry
        }
        self.input_rows.append(row_widgets)

        if data:
            division_entry.insert(0, data.get("참가부", ""))
            weight_class_entry.insert(0, data.get("체급", ""))
            count_entry.insert(0, str(data.get("인원수", "")))

    def remove_input_row(self, row_widgets):
        row_widgets['frame'].destroy()
        self.input_rows.remove(row_widgets)
        if not self.input_rows:
            self.add_input_row()

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
                data = {"참가부": row[0], "체급": row[1], "인원수": row[2]}
                self.add_input_row(data)
            
            if not self.input_rows: 
                self.add_input_row()

        except Exception as e:
            messagebox.showerror("가져오기 실패", f"엑셀 파일을 읽는 중 오류가 발생했습니다:\n{e}", parent=self)

    def download_excel_template(self):
        save_path = filedialog.asksaveasfilename(
            parent=self,
            title="엑셀 양식 저장",
            initialfile="겨루기_경기시간_계산_양식.xlsx",
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
        settings_window.title("겨루기 시간 옵션")
        settings_window.geometry("450x550")
        settings_window.transient(self)
        settings_window.grab_set()

        unique_divisions = list(set(row['division'].get() for row in self.input_rows if row['division'].get()))

        sort_order = ["초", "중", "고", "대", "일"]
        def custom_sort_key(division_name):
            for i, keyword in enumerate(sort_order):
                if keyword in division_name:
                    return (i, division_name)
            return (len(sort_order), division_name)

        divisions = sorted(unique_divisions, key=custom_sort_key)

        if not divisions:
            tk.Label(settings_window, text="\n'참가부'를 입력하고 옵션창을 열어주세요.", font=("Helvetica", 12)).pack(pady=20)
            return

        self.settings_entries = {}
        
        # --- 일괄 적용 프레임 ---
        bulk_apply_frame = tk.Frame(settings_window, pady=10)
        bulk_apply_frame.pack(fill='x', padx=15)

        tk.Label(bulk_apply_frame, text="'").pack(side="left")
        keyword_entry = tk.Entry(bulk_apply_frame, width=8)
        keyword_entry.pack(side="left")
        tk.Label(bulk_apply_frame, text="' 포함 일괄적용").pack(side="left")

        time_entry = tk.Entry(bulk_apply_frame, width=8)
        time_entry.pack(side="left", padx=(10, 2))
        tk.Label(bulk_apply_frame, text="초").pack(side="left")

        def apply_bulk_settings():
            keyword = keyword_entry.get()
            time_val = time_entry.get()
            if not time_val.isdigit():
                messagebox.showwarning("입력 오류", "시간(초)은 숫자로 입력해주세요.", parent=settings_window)
                return

            for division, var in self.settings_entries.items():
                if not keyword or keyword in division:
                    var.set(time_val)

        apply_button = tk.Button(bulk_apply_frame, text="적용", command=apply_bulk_settings)
        apply_button.pack(side="left", padx=5)

        # --- 구분선 ---
        separator = ttk.Separator(settings_window, orient='horizontal')
        separator.pack(fill='x', padx=15, pady=5)

        # --- 개별 설정 프레임 ---
        main_frame = tk.Frame(settings_window)
        main_frame.pack(padx=15, pady=15, fill="both", expand=True)
        
        tk.Label(main_frame, text="각 참가부의 경기당 소요시간(초)을 입력하세요.", font=("Helvetica", 10)).pack(pady=(0,10))

        for division in divisions:
            frame = tk.Frame(main_frame)
            frame.pack(fill="x", pady=2)
            tk.Label(frame, text=division, width=15, anchor="w").pack(side="left")
            
            var = tk.StringVar()
            entry = tk.Entry(frame, width=10, textvariable=var)
            entry.pack(side="left", padx=5)
            
            tk.Label(frame, text="초").pack(side="left")
            
            min_label = tk.Label(frame, text="", width=10, fg="gray")
            min_label.pack(side="left", padx=5)

            def update_minutes(name, index, mode, var=var, min_label=min_label):
                try:
                    seconds = float(var.get())
                    min_label.config(text=f"({seconds / 60:.1f}분)")
                except (ValueError, tk.TclError):
                    min_label.config(text="")
            
            var.trace_add("write", update_minutes)
            self.settings_entries[division] = var

        button_frame = tk.Frame(settings_window)
        button_frame.pack(fill='x', side='bottom', pady=10, padx=15)
        tk.Button(button_frame, text="저장", command=lambda: self.save_settings(settings_window)).pack(side="right")

        self.load_settings()

    def load_settings(self):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                settings = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            settings = {}

        for division, var in self.settings_entries.items():
            var.set(settings.get(division, "450"))

    def save_settings(self, window):
        settings = {}
        for division, var in self.settings_entries.items():
            settings[division] = var.get()

        try:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            window.destroy()
        except Exception as e:
            messagebox.showerror("저장 실패", f"설정을 저장하는 중 오류가 발생했습니다:\n{e}", parent=window)

