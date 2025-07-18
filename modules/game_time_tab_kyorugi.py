import tkinter as tk
from tkinter import ttk, messagebox, filedialog, font
import json
import shutil
import os
import openpyxl
from datetime import datetime, timedelta

from common.constants import KYORUGI_SETTINGS_FILE as SETTINGS_FILE
from common.constants import KYORUGI_TEMPLATE_PATH as TEMPLATE_PATH
from utils.file_operations import download_template_file

class KyorugiTab(ttk.Frame):
    def __init__(self, notebook, parent_app):
        super().__init__(notebook)
        self.parent_app = parent_app
        self.input_rows = []
        self.settings_entries = {}
        self.color_palette = ["#ADD8E6", "#90EE90", "#FFFFE0", "#FFDAB9", "#E6E6FA", "#B0E0E6", "#FFE4E1", "#D8BFD8", "#F5DEB3", "#C0C0C0"]
        self.color_index = 0
        self.text_color_map = {}
        self.create_widgets()
        self.populate_default_rows()

    def _generate_round_options(self, headcount):
        options = []
        
        # Generate powers of 2 rounds based on headcount exceeding half the round size
        for power in [1024, 512, 256, 128, 64, 32, 16, 8]:
            if headcount > (power / 2):
                options.append(f"{power}강")
        
        # Specific conditions for semi-final and final
        if headcount > 2:
            options.append("준결승")
        if headcount >= 2:
            options.append("결승")
            
        return options

    def create_widgets(self):
        main_paned_window = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_paned_window.pack(fill=tk.BOTH, expand=True)

        left_frame = ttk.Frame(main_paned_window, width=550)
        right_frame = ttk.Frame(main_paned_window, width=350)
        main_paned_window.add(left_frame, weight=11)
        main_paned_window.add(right_frame, weight=7)

        control_frame = tk.Frame(left_frame)
        control_frame.pack(fill='x', padx=10, pady=(10,0))

        add_button = tk.Button(control_frame, text="+ 행추가", command=self.add_input_row)
        add_button.pack(side="left", padx=(0, 5))
        import_button = tk.Button(control_frame, bg="#4CAF50", fg="white", text="엑셀로 가져오기", command=self.import_from_excel)
        import_button.pack(side="left", padx=5)
        download_button = tk.Button(control_frame, text="엑셀 양식 다운로드", command=self.download_excel_template)
        download_button.pack(side="left", padx=5)
        settings_button = tk.Button(control_frame, text="⚙️ 옵션", command=self.open_kyorugi_settings)
        settings_button.pack(side="right")

        filter_frame = tk.Frame(left_frame)
        filter_frame.pack(fill='x', padx=10, pady=(10,10))

        self.filter_comboboxes = {}

        division_filter_values = [""] # Will be populated dynamically
        self.division_filter_combo = ttk.Combobox(filter_frame, values=division_filter_values, width=15, state="readonly")
        self.division_filter_combo.set("참가부 필터")
        self.division_filter_combo.pack(side="left", padx=2)
        self.division_filter_combo.bind("<<ComboboxSelected>>", self._apply_filters)
        self.filter_comboboxes["division"] = self.division_filter_combo

        weight_class_filter_values = [""] # Will be populated dynamically
        self.weight_class_filter_combo = ttk.Combobox(filter_frame, values=weight_class_filter_values, width=15, state="readonly")
        self.weight_class_filter_combo.set("체급 필터")
        self.weight_class_filter_combo.pack(side="left", padx=2)
        self.weight_class_filter_combo.bind("<<ComboboxSelected>>", self._apply_filters)
        self.filter_comboboxes["weight_class"] = self.weight_class_filter_combo

        clear_filter_button = tk.Button(filter_frame, text="필터 초기화", command=self._clear_filters)
        clear_filter_button.pack(side="left", padx=5)

        input_grid_frame = tk.Frame(left_frame)
        input_grid_frame.pack(expand=True, fill="both", padx=10, pady=10)

        header_frame = tk.Frame(input_grid_frame)
        header_frame.pack(fill='x', pady=(5, 5))
        self.header_check_var = tk.IntVar()
        header_check = tk.Checkbutton(header_frame, variable=self.header_check_var, command=self.toggle_all_checks)
        header_check.pack(side="left", padx=2, anchor='w')
        tk.Label(header_frame, text="참가부", width=15, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="체급", width=13, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="인원수", width=8, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="시작 강수", width=15, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="종료 강수", width=15, anchor='w').pack(side="left", padx=2)

        self.canvas = tk.Canvas(input_grid_frame)
        scrollbar = tk.Scrollbar(input_grid_frame, orient="vertical", command=self.canvas.yview)
        self.rows_container = tk.Frame(self.canvas)
        self.rows_container.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.rows_container, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.populate_default_rows()

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
        self.result_text.tag_configure("bold", font=("Helvetica", 10, "bold"))

        # Footer
        footer_frame = tk.Frame(right_frame)
        footer_frame.pack(side=tk.BOTTOM, pady=5)

        footer_font = font.Font(family="Helvetica", size=9)
        footer_label = tk.Label(footer_frame, text="Copyright (c) FEELJAE-WON. All rights reserved.", font=footer_font, fg="gray")
        footer_label.pack(side=tk.LEFT, padx=5)

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
                start_round = row['start_round_var'].get()
                end_round = row['end_round_var'].get()
                
                if not division or headcount == 0:
                    continue

                time_per_match = int(settings.get(division, 450)) # 기본값 450초
                
                num_matches = self._get_matches_for_round_range(headcount, start_round, end_round)
                
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

        result_str = "==================== 코트 적용 소요시간 ====================\n\n"
        result_str += f"총 예상 소요시간: {format_time(total_duration_seconds)}\n"
        result_str += "\n[겨루기] - " + str(court_count) + " 코트 기준\n"
        result_str += f"  총 소요시간: {format_time(kyorugi_duration_per_court)}\n\n"
        
        result_str += "============================================================\n\n"
        result_str += f"시작 시간: {start_time.strftime('%H:%M')}\n"
        result_str += f"예상 종료 시간: {end_time.strftime('%H:%M')}\n\n"
        result_str += "============================================================\n"

        # 참가부별 코트 반영 소요시간 및 게임 수
        division_data = {}
        for row in rows_to_process:
            division = row['division'].get()
            weight_class = row['weight_class'].get()
            headcount = int(row['count'].get() or 0)

            if not division or headcount == 0:
                continue

            time_per_match = int(settings.get(division, 450))
            start_round = row['start_round_var'].get()
            end_round = row['end_round_var'].get()
            num_matches = self._get_matches_for_round_range(headcount, start_round, end_round)
            row_total_seconds = num_matches * time_per_match

            if division not in division_data:
                division_data[division] = {"total_seconds": 0, "total_games": 0}
            division_data[division]["total_seconds"] += row_total_seconds
            division_data[division]["total_games"] += num_matches

        if division_data:
            result_str += "\n========== 참가부별 코트 반영 소요시간 및 게임 수 ==========\n\n"
            for division, data in division_data.items():
                adjusted_seconds = data["total_seconds"] / court_count if court_count > 0 else 0
                result_str += f"  {division}: {format_time(adjusted_seconds)} (총 {data["total_games"]} 게임)\n"

        applied_settings_summary = {}
        for row in rows_to_process:
            division = row['division'].get()
            if division and division not in applied_settings_summary:
                time_per_match = int(settings.get(division, 450))
                applied_settings_summary[division] = time_per_match

        if applied_settings_summary:
            result_str += "\n============== 적용된 참가부별 경기 시간 설정 ==============\n\n"
            for division, time_in_seconds in applied_settings_summary.items():
                minutes = time_in_seconds / 60
                result_str += f"  {division}: {time_in_seconds}초 ({minutes:.1f}분)\n"
            result_str += "\n============================================================\n"

        self.result_text.config(state="normal")
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, result_str)

        # Apply bold tags
        self.result_text.tag_add("bold", "3.0", "3.end") # 총 예상 소요시간
        self.result_text.tag_add("bold", "5.0", "5.end") # 시작 시간
        self.result_text.tag_add("bold", "6.0", "6.end") # 예상 종료 시간

        self.result_text.config(state="disabled")

    def _get_matches_for_round_range(self, headcount, start_round_str, end_round_str):
        def _extract_round_size(round_str):
            if round_str == "결승":
                return 2
            elif round_str == "준결승":
                return 4
            try:
                return int(round_str.replace("강", ""))
            except ValueError:
                return 0

        if headcount <= 1:
            return 0

        start_round_size = _extract_round_size(start_round_str)
        end_round_size = _extract_round_size(end_round_str)

        # Ensure start_round_size is at least as large as end_round_size for a valid range
        if start_round_size < end_round_size:
            start_round_size = end_round_size

        total_matches = 0
        
        # Find the smallest power of 2 that is greater than or equal to headcount
        initial_bracket_size = 2
        while initial_bracket_size < headcount:
            initial_bracket_size *= 2
        
        # Iterate through bracket sizes from the initial_bracket_size down to 2 (결승)
        current_bracket_power = initial_bracket_size

        while current_bracket_power >= 2:
            matches_in_this_round = 0
            
            if current_bracket_power == initial_bracket_size:
                # This is the first round where byes might occur
                # Number of players who actually play in this round
                players_playing_this_round = headcount - (current_bracket_power - headcount)
                matches_in_this_round = players_playing_this_round // 2
            else:
                # For subsequent rounds, the number of participants is exactly the bracket size
                matches_in_this_round = current_bracket_power // 2
            
            # Check if this round (bracket size) is within the user's selected range
            if current_bracket_power <= start_round_size and current_bracket_power >= end_round_size:
                total_matches += matches_in_this_round
            
            current_bracket_power //= 2 # Move to the next smaller bracket (e.g., 64 -> 32)

        return total_matches

    def _update_division_entry_color(self, entry):
        text = entry.get()
        if text:
            color = self._get_color_for_text(text)
            entry.config(bg=color, fg="black")
        else:
            entry.config(bg="white", fg="black")

    def _update_count_entry_style(self, entry, var, row_widgets):
        try:
            count = int(var.get())
            if count < 4:
                entry.config(bg="red", fg="white")
            else:
                entry.config(bg="white", fg="black")
        except ValueError:
            entry.config(bg="white", fg="black") # Reset if not a valid number
        
        # Update round options whenever headcount changes
        if row_widgets:
            self._update_row_round_options(row_widgets)

    def _get_color_for_text(self, text):
        if text not in self.text_color_map:
            self.text_color_map[text] = self.color_palette[self.color_index]
            self.color_index = (self.color_index + 1) % len(self.color_palette)
        return self.text_color_map[text]

    def _update_filter_options(self):
        divisions = sorted(list(set(row['division'].get() for row in self.input_rows if row['division'].get())))
        weight_classes = sorted(list(set(row['weight_class'].get() for row in self.input_rows if row['weight_class'].get())))

        self.division_filter_combo['values'] = [""] + divisions
        self.weight_class_filter_combo['values'] = [""] + weight_classes

    def _apply_filters(self, event=None):
        selected_division = self.division_filter_combo.get()
        selected_weight_class = self.weight_class_filter_combo.get()

        for row_widgets in self.input_rows:
            division_match = (selected_division == "" or selected_division == "참가부 필터" or row_widgets['division'].get() == selected_division)
            weight_class_match = (selected_weight_class == "" or selected_weight_class == "체급 필터" or row_widgets['weight_class'].get() == selected_weight_class)

            if division_match and weight_class_match:
                row_widgets['frame'].pack(fill='x', pady=2, anchor='w')
            else:
                row_widgets['frame'].pack_forget()
        self.canvas.yview_moveto(0) # Scroll to top

    def _clear_filters(self):
        self.division_filter_combo.set("참가부 필터")
        self.weight_class_filter_combo.set("체급 필터")
        self._apply_filters()

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
        self._update_filter_options()
        self._apply_filters()

    def toggle_all_checks(self):
        is_checked = self.header_check_var.get()
        for row_widgets in self.input_rows:
            # Only toggle if the row is currently visible (packed)
            if row_widgets['frame'].winfo_ismapped():
                row_widgets['check_var'].set(is_checked)

    def add_input_row(self, data=None):
        row_frame = tk.Frame(self.rows_container)
        row_frame.pack(fill='x', pady=2, anchor='w')

        check_var = tk.IntVar()
        check = tk.Checkbutton(row_frame, variable=check_var)
        check.pack(side="left", padx=2)

        division_var = tk.StringVar()
        division_entry = tk.Entry(row_frame, width=15, textvariable=division_var)
        division_entry.pack(side="left", padx=2)
        division_var.trace_add("write", lambda name, index, mode, entry=division_entry: self._update_division_entry_color(entry))

        weight_class_entry = tk.Entry(row_frame, width=13, bg="white", fg="black") # Set default color
        weight_class_entry.pack(side="left", padx=2)

        count_var = tk.StringVar()
        count_entry = tk.Entry(row_frame, width=8, textvariable=count_var)
        count_entry.pack(side="left", padx=2)
        count_var.trace_add("write", lambda name, index, mode, entry=count_entry, var=count_var, row_widgets=None: self._update_count_entry_style(entry, var, row_widgets))

        tk.Label(row_frame, text="(").pack(side="left")
        start_round_var = tk.StringVar()
        start_round_combo = ttk.Combobox(row_frame, textvariable=start_round_var, values=[], width=10, state="readonly")
        start_round_combo.pack(side="left", padx=2)

        tk.Label(row_frame, text="~").pack(side="left")
        end_round_var = tk.StringVar()
        end_round_combo = ttk.Combobox(row_frame, textvariable=end_round_var, values=[], width=10, state="readonly")
        end_round_combo.pack(side="left", padx=2)
        tk.Label(row_frame, text=")").pack(side="left")
        
        delete_button = tk.Button(row_frame, text="-", command=lambda: self.remove_input_row(row_widgets))
        delete_button.pack(side="left", padx=2)

        row_widgets = {
            'frame': row_frame, 'check_var': check_var, 'division': division_entry, 
            'weight_class': weight_class_entry, 'count': count_entry,
            'division_var': division_var, 'count_var': count_var,
            'start_round_var': start_round_var, 'end_round_var': end_round_var,
            'start_round_combo': start_round_combo, 'end_round_combo': end_round_combo
        }
        self.input_rows.append(row_widgets)

        # Pass row_widgets to the trace after it's fully defined
        count_var.trace_add("write", lambda name, index, mode, entry=count_entry, var=count_var: self._update_count_entry_style(entry, var, row_widgets))

        if data:
            division_entry.insert(0, data.get("참가부", ""))
            weight_class_entry.insert(0, data.get("체급", ""))
            count_entry.insert(0, str(data.get("인원수", "")))
            self._update_count_entry_style(count_entry, count_var, row_widgets)
            self._update_row_round_options(row_widgets, data.get("시작강수"), data.get("종료강수"))
        else:
            # For new rows, ensure initial round options are set
            self._update_row_round_options(row_widgets)

        self._update_filter_options()
        self._apply_filters()

    def _update_row_round_options(self, row_widgets, initial_start_round=None, initial_end_round=None):
        try:
            headcount = int(row_widgets['count_var'].get() or 0)
        except ValueError:
            headcount = 0

        options = self._generate_round_options(headcount)
        
        row_widgets['start_round_combo']['values'] = options
        row_widgets['end_round_combo']['values'] = options

        # Set start round
        if initial_start_round and initial_start_round in options:
            row_widgets['start_round_var'].set(initial_start_round)
        elif options:
            row_widgets['start_round_var'].set(options[0])
        else:
            row_widgets['start_round_var'].set("")

        # Set end round
        if initial_end_round and initial_end_round in options:
            row_widgets['end_round_var'].set(initial_end_round)
        elif options:
            row_widgets['end_round_var'].set(options[-1])
        else:
            row_widgets['end_round_var'].set("")

    def remove_input_row(self, row_widgets):
        row_widgets['frame'].destroy()
        self.input_rows.remove(row_widgets)
        if not self.input_rows:
            self.add_input_row()
        self._update_filter_options()
        self._apply_filters()

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
                # Assuming Excel might have '시작강수' and '종료강수' columns
                if len(row) > 3: data["시작강수"] = row[3]
                if len(row) > 4: data["종료강수"] = row[4]
                self.add_input_row(data)
            
            if not self.input_rows: 
                self.add_input_row()
            self._update_filter_options()
            self._apply_filters()

        except Exception as e:
            messagebox.showerror("가져오기 실패", f"엑셀 파일을 읽는 중 오류가 발생했습니다:\n{e}", parent=self)

    def download_excel_template(self):
        download_template_file(TEMPLATE_PATH, "겨루기_경기시간_계산_양식.xlsx", [("Excel files", "*.xlsx"), ("All files", "*.* ")])

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

