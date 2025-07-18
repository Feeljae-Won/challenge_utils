import tkinter as tk
from tkinter import ttk, font, messagebox, filedialog
import json
import shutil
import os
import openpyxl
from datetime import datetime, timedelta

from common.constants import POOMSAE_SETTINGS_FILE as SETTINGS_FILE
from common.constants import POOMSAE_TEMPLATE_PATH as TEMPLATE_PATH
from utils.file_operations import download_template_file

DEFAULT_SETTINGS = {
    "individual": {
        "초등부": "210", "중등부": "210", "고등부": "210", "대학부": "240", "일반부": "240"
    },
    "team": {
        "초등부": "390", "중등부": "420", "고등부": "390", "대학부": "450", "일반부": "420"
    },
    "freestyle": {
        "개인": "140", "복식": "140", "단체": "140"
    }
}

class PoomsaeTab(ttk.Frame):
    def __init__(self, notebook, parent_app):
        super().__init__(notebook)
        self.parent_app = parent_app # Reference to the main GameTimeCalculator app
        self.input_rows = []
        self.rows_container = None # Initialize rows_container
        self.create_widgets()

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
        settings_button = tk.Button(control_frame, text="⚙️ 옵션", command=self.open_poomsae_settings)
        settings_button.pack(side="right")

        filter_frame = tk.Frame(left_frame)
        filter_frame.pack(fill='x', padx=10, pady=(10,10))

        self.filter_comboboxes = {}

        event_filter_values = ["", "개인전", "복식전", "단체전", "개인전(자유품새)", "복식전(자유품새)", "단체전(자유품새)"]
        self.event_filter_combo = ttk.Combobox(filter_frame, values=event_filter_values, width=15, state="readonly")
        self.event_filter_combo.set("종목 필터")
        self.event_filter_combo.pack(side="left", padx=2)
        self.event_filter_combo.bind("<<ComboboxSelected>>", self._apply_filters)
        self.filter_comboboxes["event"] = self.event_filter_combo

        division_filter_values = ["", "초등부", "중등부", "고등부", "대학부", "일반부"]
        self.division_filter_combo = ttk.Combobox(filter_frame, values=division_filter_values, width=15, state="readonly")
        self.division_filter_combo.set("참가부 필터")
        self.division_filter_combo.pack(side="left", padx=2)
        self.division_filter_combo.bind("<<ComboboxSelected>>", self._apply_filters)
        self.filter_comboboxes["division"] = self.division_filter_combo

        gender_filter_values = ["", "남자", "여자", "혼성"]
        self.gender_filter_combo = ttk.Combobox(filter_frame, values=gender_filter_values, width=8, state="readonly")
        self.gender_filter_combo.set("성별 필터")
        self.gender_filter_combo.pack(side="left", padx=2)
        self.gender_filter_combo.bind("<<ComboboxSelected>>", self._apply_filters)
        self.filter_comboboxes["gender"] = self.gender_filter_combo

        # Add a clear filter button
        clear_filter_button = tk.Button(filter_frame, text="필터 초기화", command=self._clear_filters)
        clear_filter_button.pack(side="left", padx=5)

        input_grid_frame = tk.Frame(left_frame)
        input_grid_frame.pack(expand=True, fill="both", padx=10, pady=10)

        header_frame = tk.Frame(input_grid_frame)
        header_frame.pack(fill='x', pady=(5, 5))
        self.header_check_var = tk.IntVar()
        header_check = tk.Checkbutton(header_frame, variable=self.header_check_var, command=self.toggle_all_checks)
        header_check.pack(side="left", padx=2, anchor='w')
        tk.Label(header_frame, text="종목", width=18, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="참가부", width=18, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="세부부별", width=15, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="성별", width=10, anchor='w').pack(side="left", padx=2)
        tk.Label(header_frame, text="인원수", width=15, anchor='w').pack(side="left", padx=2)

        
        canvas = tk.Canvas(input_grid_frame)
        scrollbar = tk.Scrollbar(input_grid_frame, orient="vertical", command=canvas.yview)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.rows_container = tk.Frame(canvas)
        self.rows_container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.rows_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
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
        tk.Label(court_frame, text="공인 코트수:").pack(side="left", padx=(0,5))
        self.gongin_court_entry = tk.Entry(court_frame, width=5)
        self.gongin_court_entry.insert(0, "4")
        self.gongin_court_entry.pack(side="left", padx=(0, 10))
        tk.Label(court_frame, text="자유 코트수:").pack(side="left", padx=(0,5))
        self.jayu_court_entry = tk.Entry(court_frame, width=5)
        self.jayu_court_entry.insert(0, "2")
        self.jayu_court_entry.pack(side="left")

        self.freestyle_simultaneous_var = tk.IntVar(value=0) # Default to unchecked
        freestyle_simultaneous_check = tk.Checkbutton(court_frame, text="자유품새 동시진행", variable=self.freestyle_simultaneous_var)
        freestyle_simultaneous_check.pack(side="left", padx=5)

        freestyle_calc_options_frame = ttk.LabelFrame(results_labelframe, text="자유품새 계산")
        freestyle_calc_options_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.prelim_var = tk.IntVar(value=1) # Default to checked
        self.main_var = tk.IntVar(value=1) # Default to checked
        self.final_var = tk.IntVar(value=1) # Default to checked

        ttk.Checkbutton(freestyle_calc_options_frame, text="예선", variable=self.prelim_var).pack(side="left", padx=5, pady=5)
        ttk.Checkbutton(freestyle_calc_options_frame, text="본선", variable=self.main_var).pack(side="left", padx=5, pady=5)
        ttk.Checkbutton(freestyle_calc_options_frame, text="결선", variable=self.final_var).pack(side="left", padx=5, pady=5)

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
            settings = DEFAULT_SETTINGS

        try:
            start_time = datetime.strptime(self.start_time_var.get(), "%H:%M")
            gongin_courts = int(self.gongin_court_entry.get())
            jayu_courts = int(self.jayu_court_entry.get())
            if gongin_courts <= 0 or jayu_courts <= 0:
                raise ValueError("코트 수는 0보다 커야 합니다.")
        except ValueError as e:
            messagebox.showerror("입력 오류", f"시작 시간 또는 코트 수 입력이 잘못되었습니다.\n{e}", parent=self)
            return

        selected_rows = [row for row in self.input_rows if row['check_var'].get() == 1]
        rows_to_process = selected_rows if selected_rows else self.input_rows

        total_seconds_gongin_raw = 0
        total_seconds_jayu_raw = 0

        sub_totals = {
            "공인품새": {"개인전": {"time": 0, "games": 0}, "복식전": {"time": 0, "games": 0}, "단체전": {"time": 0, "games": 0}},
            "자유품새": {"개인전": {"time": 0, "games": 0}, "복식전": {"time": 0, "games": 0}, "단체전": {"time": 0, "games": 0}}
        }

        for row in rows_to_process:
            try:
                event_input = row['event'].get() # e.g., "개인전", "개인전(자유품새)"
                division_input = row['division'].get() # e.g., "초등부", "일반부"
                category = row['class'].get()
                gender = row['gender'].get()
                original_headcount = int(row['count'].get() or 0)
                calculated_headcount = original_headcount # Initialize with original headcount

                # Determine the actual event type (공인품새 or 자유품새) for calculation and sub_totals key
                actual_event_type = "공인품새"
                if "자유품새" in event_input:
                    actual_event_type = "자유품새"

                # Determine the actual division type (개인전, 복식전, 단체전) from event_input
                actual_division_type = event_input.replace("(자유품새)", "").strip()

                # Validate division_input (참가부) against expected categories
                valid_categories = ["초등부", "중등부", "고등부", "대학부", "일반부"]
                if division_input not in valid_categories:
                    messagebox.showerror("데이터 오류", f"참가부 입력이 잘못되었습니다. {', '.join(valid_categories)} 중 하나여야 합니다.\n잘못된 값: {division_input}", parent=self)
                    return # Stop calculation if invalid data is found

                # Apply headcount adjustments based on actual_event_type and actual_division_type
                if actual_event_type == "공인품새":
                    if actual_division_type == "개인전":
                        calculated_headcount = original_headcount - 1
                    elif actual_division_type == "복식전":
                        calculated_headcount = (original_headcount / 2) - 1
                    elif actual_division_type == "단체전":
                        calculated_headcount = (original_headcount / 3) - 1
                else: # 자유품새
                    if actual_division_type == "개인전":
                        if original_headcount > 22:
                            participants = original_headcount
                            num_prelim_groups = 2
                            while participants / num_prelim_groups > 11.5 and num_prelim_groups % 2 == 0:
                                num_prelim_groups += 2

                            if num_prelim_groups < 2:
                                num_prelim_groups = 2
                            if num_prelim_groups % 2 != 0:
                                num_prelim_groups += 1

                            base_prelim_group_size = participants // num_prelim_groups
                            remainder_prelim = participants % num_prelim_groups

                            prelim_group_sizes = []
                            for g in range(num_prelim_groups):
                                size = base_prelim_group_size
                                if g < remainder_prelim:
                                    size += 1
                                prelim_group_sizes.append(size)

                            # Calculate advancement to main round based on first group size
                            first_group_size = prelim_group_sizes[0] if prelim_group_sizes else 0
                            advancement_per_group = (first_group_size + 1) // 2 # Round up
                            total_main_round_participants_for_next_stage = advancement_per_group * num_prelim_groups

                            # Now, calculate main round groups based on total_main_round_participants_for_next_stage
                            num_main_groups = 2
                            while total_main_round_participants_for_next_stage / num_main_groups > 11.5 and num_main_groups % 2 == 0:
                                num_main_groups += 2

                            if num_main_groups < 2:
                                num_main_groups = 2
                            if num_main_groups % 2 != 0:
                                num_main_groups += 1

                            base_main_group_size = total_main_round_participants_for_next_stage // num_main_groups
                            remainder_main = total_main_round_participants_for_next_stage % num_main_groups

                            main_group_sizes = []
                            for g in range(num_main_groups):
                                size = base_main_group_size
                                if g < remainder_main:
                                    size += 1
                                main_group_sizes.append(size)

                            calculated_headcount = 0
                            if self.prelim_var.get():
                                calculated_headcount += sum(prelim_group_sizes)
                            if self.main_var.get():
                                calculated_headcount += sum(main_group_sizes)
                            if self.final_var.get():
                                calculated_headcount += 8 # 결선은 항상 8명

                        elif 12 <= original_headcount <= 21:
                            calculated_headcount = 0
                            if self.main_var.get():
                                calculated_headcount += original_headcount
                            if self.final_var.get():
                                calculated_headcount += 8 # 결선은 항상 8명
                        # else: original_headcount <= 11, calculated_headcount remains original_headcount
                        # If only final is checked for <= 11 participants, it should be 8
                        if original_headcount <= 11 and self.final_var.get() and not self.prelim_var.get() and not self.main_var.get():
                            calculated_headcount = 8
                        elif original_headcount <= 11 and not self.final_var.get() and not self.prelim_var.get() and not self.main_var.get():
                            calculated_headcount = 0 # If nothing is checked, and participants <= 11, then 0
                        elif original_headcount <= 11 and self.final_var.get() and (self.prelim_var.get() or self.main_var.get()):
                            calculated_headcount = original_headcount + 8 # If final is checked and other rounds are also checked, add 8 for final
                        elif original_headcount <= 11 and (self.prelim_var.get() or self.main_var.get()):
                            calculated_headcount = original_headcount # If only prelim/main are checked, use original headcount
                    elif actual_division_type == "복식전":
                        calculated_headcount = original_headcount / 2
                    elif actual_division_type == "단체전":
                        calculated_headcount = original_headcount / 5

                time_per_game = 0
                if actual_event_type == "공인품새":
                    if actual_division_type == "개인전":
                        time_per_game = int(settings['individual'].get(division_input, 0)) # Use division_input as category
                    elif actual_division_type == "복식전" or actual_division_type == "단체전":
                        time_per_game = int(settings['team'].get(division_input, 0)) # Use division_input as category
                else: # 자유품새
                    freestyle_map = {"개인전": "개인", "복식전": "복식", "단체전": "단체"}
                    time_per_game = int(settings['freestyle'].get(freestyle_map.get(actual_division_type), 0))
                
                row_total_seconds = calculated_headcount * time_per_game
                
                # Accumulate raw total seconds for each poomsae type
                if actual_event_type == "공인품새":
                    total_seconds_gongin_raw += row_total_seconds
                else:
                    total_seconds_jayu_raw += row_total_seconds

                # Accumulate raw sub-totals for each division
                # Use actual_event_type for the first level key
                sub_totals[actual_event_type][actual_division_type]["time"] += row_total_seconds
                sub_totals[actual_event_type][actual_division_type]["games"] += calculated_headcount # Store calculated games

            except (ValueError, KeyError) as e:
                messagebox.showerror("데이터 오류", f"입력 데이터에 오류가 있습니다. 확인해주세요.\n종목: {event_input}, 참가부: {division_input}, 세부부별: {category}, 성별: {gender}\n오류: {e}", parent=self)
                return

        # Calculate court-applied durations for each poomsae type
        gongin_duration_per_court = total_seconds_gongin_raw / gongin_courts if gongin_courts > 0 else 0
        
        jayu_duration_per_court = total_seconds_jayu_raw
        if self.freestyle_simultaneous_var.get() == 1: # If checkbox is checked, divide by court count
            jayu_duration_per_court = total_seconds_jayu_raw / jayu_courts if jayu_courts > 0 else 0

        # Total estimated time is the sum of court-applied times for each poomsae type
        total_duration_seconds = gongin_duration_per_court + jayu_duration_per_court
        end_time = start_time + timedelta(seconds=total_duration_seconds)

        # --- Display Results ---
        def format_time(seconds):
            m, s = divmod(seconds, 60)
            h, m = divmod(m, 60)
            return f"{int(h)}시간 {int(m)}분 {int(s)}초"

        def format_subtotal_with_games(seconds, games_count, court_divisor=1):
            # Apply court division here for individual division times
            effective_seconds = seconds / court_divisor if court_divisor > 0 else 0
            return f"{format_time(effective_seconds)} (총 {int(games_count)} 게임)"

        result_str = "==================== 코트 적용 소요시간 ====================\n"
        result_str += f"\n총 예상 소요시간: {format_time(total_duration_seconds)}\n"
        result_str += f"\n시작 시간: {start_time.strftime('%H:%M')}\n"
        end_time_str = end_time.strftime('%H:%M')
        if end_time.day != start_time.day:
            end_time_str += " (+1)"
        result_str += f"\n예상 종료 시간: {end_time_str}\n"
        result_str += "\n============================================================\n\n"
        
        result_str += "[공인품새] - " + str(gongin_courts) + " 코트 기준\n\n"
        result_str += f"  개인전 소요시간: {format_subtotal_with_games(sub_totals['공인품새']['개인전']['time'], sub_totals['공인품새']['개인전']['games'], gongin_courts)}\n"
        result_str += f"  복식전 소요시간: {format_subtotal_with_games(sub_totals['공인품새']['복식전']['time'], sub_totals['공인품새']['복식전']['games'], gongin_courts)}\n"
        result_str += f"  단체전 소요시간: {format_subtotal_with_games(sub_totals['공인품새']['단체전']['time'], sub_totals['공인품새']['단체전']['games'], gongin_courts)}\n\n"
        total_gongin_games = sub_totals['공인품새']['개인전']['games'] + \
                             sub_totals['공인품새']['복식전']['games'] + \
                             sub_totals['공인품새']['단체전']['games']
        result_str += f"  공인품새 총 소요시간: {format_time(gongin_duration_per_court)} (총 {int(total_gongin_games)} 게임)\n"

        freestyle_simultaneous_status = "적용" if self.freestyle_simultaneous_var.get() == 1 else "미적용"
        total_jayu_games = sub_totals['자유품새']['개인전']['games'] + \
                           sub_totals['자유품새']['복식전']['games'] + \
                           sub_totals['자유품새']['단체전']['games']
        result_str += f"\n[자유품새] - {jayu_courts} 코트 기준 (동시진행 {freestyle_simultaneous_status})\n\n"
        result_str += f"  개인전 소요시간: {format_subtotal_with_games(sub_totals['자유품새']['개인전']['time'], sub_totals['자유품새']['개인전']['games'], jayu_courts if self.freestyle_simultaneous_var.get() == 1 else 1)}\n"
        result_str += f"  복식전 소요시간: {format_subtotal_with_games(sub_totals['자유품새']['복식전']['time'], sub_totals['자유품새']['복식전']['games'], jayu_courts if self.freestyle_simultaneous_var.get() == 1 else 1)}\n"
        result_str += f"  단체전 소요시간: {format_subtotal_with_games(sub_totals['자유품새']['단체전']['time'], sub_totals['자유품새']['단체전']['games'], jayu_courts if self.freestyle_simultaneous_var.get() == 1 else 1)}\n\n"
        result_str += f"  자유품새 총 소요시간: {format_time(jayu_duration_per_court)} (총 {int(total_jayu_games)} 게임)\n"

        result_str += "\n============================================================\n"
        
        result_str += "\n ** 설명 **\n\n"
        result_str += "   - 공인품새\n\n"
        result_str += "       * 개인전의 경우 각 행마다 (인원수 - 1) * 소요시간으로 계산\n"
        result_str += "       * 복/식단체의 경우 각 행마다 {(인원수 / 2) - 1} * 소요시간으로 계산\n\n"
        result_str += "   - 자유품새\n\n"
        result_str += "       * 개인전의 경우 각 행마다 (인원수 - 1) * 소요시간으로 계산\n"
        result_str += "       * 복/식단체의 경우 각 행마다 {(인원수 / 2 또는 5) - 1} * 소요시간으로 계산\n"
        result_str += "       * 22명(팀) 이상 참가할 경우 [{인원수 + (인원수 * 0.5)} + 8] * 소요시간 으로 \n"
        result_str += "         예선 - 본선 - 결선을 계산\n"
        result_str += "       * 12명(팀) 이상 21명(팀) 이하 참가할 경우 (인원수 + 8) * 소요시간으로 \n"
        result_str += "         본선 - 결선을 계산\n"
        result_str += "       * 11명(팀) 이하일 경우 결선으로 계산\n"
        

        self.result_text.config(state="normal")
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, result_str)

        # Apply bold tags
        self.result_text.tag_add("bold", "3.0", "3.end") # 총 예상 소요시간
        self.result_text.tag_add("bold", "5.0", "5.end") # 시작 시간
        self.result_text.tag_add("bold", "7.0", "7.end") # 예상 종료 시간

        # Find and bold "공인품새 총 소요시간" and "자유품새 총 소요시간"
        # These lines might shift, so find them dynamically
        gongin_total_time_start = self.result_text.search("공인품새 총 소요시간", "1.0", tk.END)
        if gongin_total_time_start:
            gongin_total_time_end = self.result_text.search("\n", gongin_total_time_start, tk.END)
            if gongin_total_time_end:
                self.result_text.tag_add("bold", gongin_total_time_start, gongin_total_time_end)

        jayu_total_time_start = self.result_text.search("자유품새 총 소요시간", "1.0", tk.END)
        if jayu_total_time_start:
            jayu_total_time_end = self.result_text.search("\n", jayu_total_time_start, tk.END)
            if jayu_total_time_end:
                self.result_text.tag_add("bold", jayu_total_time_start, jayu_total_time_end)

        self.result_text.config(state="disabled")

    def set_current_time(self):
        now = datetime.now()
        current_time = now.strftime("%H:%M")
        self.start_time_var.set(current_time)

    def populate_default_rows(self):
        for i in range(len(self.input_rows) - 1, -1, -1):
            self.input_rows[i]['frame'].destroy()
            self.input_rows.pop(i)

        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                settings = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            settings = DEFAULT_SETTINGS

        all_rows_data = []

        # 공인품새
        for category in ["초등부", "중등부", "고등부", "대학부", "일반부"]:
            all_rows_data.append({"종목": "개인전", "참가부": category, "세부부별": "", "성별": "", "인원수": ""})
            all_rows_data.append({"종목": "복식전", "참가부": category, "세부부별": "", "성별": "", "인원수": ""})
            all_rows_data.append({"종목": "단체전", "참가부": category, "세부부별": "", "성별": "", "인원수": ""})

        # 자유품새
        for category in ["초등부", "중등부", "고등부", "대학부", "일반부"]:
            all_rows_data.append({"종목": "개인전(자유품새)", "참가부": category, "세부부별": "", "성별": "", "인원수": ""})
            all_rows_data.append({"종목": "복식전(자유품새)", "참가부": category, "세부부별": "", "성별": "", "인원수": ""})
            all_rows_data.append({"종목": "단체전(자유품새)", "참가부": category, "세부부별": "", "성별": "", "인원수": ""})

        # Define custom sort order for '종목'
        event_order = {"개인전": 1, "복식전": 2, "단체전": 3, "개인전(자유품새)": 4, "복식전(자유품새)": 5, "단체전(자유품새)": 6}
        # Define custom sort order for '참가부' (division)
        division_order = {"초등부": 1, "중등부": 2, "고등부": 3, "대학부": 4, "일반부": 5}

        # Sort the rows
        all_rows_data.sort(key=lambda x: (
            event_order.get(x["종목"], 99), 
            division_order.get(x["참가부"], 99)
        ))

        for data in all_rows_data:
            self.add_input_row(data)

        # Ensure there are at least 10 rows
        while len(self.input_rows) < 10:
            self.add_input_row()
        self._clear_filters() # Apply filters after populating rows

    def _apply_filters(self, event=None):
        selected_event = self.event_filter_combo.get()
        selected_division = self.division_filter_combo.get()
        selected_gender = self.gender_filter_combo.get()

        for row_widgets in self.input_rows:
            event_match = (selected_event == "" or selected_event == "종목 필터" or row_widgets['event'].get() == selected_event)
            division_match = (selected_division == "" or selected_division == "참가부 필터" or row_widgets['division'].get() == selected_division)
            gender_match = (selected_gender == "" or selected_gender == "성별 필터" or row_widgets['gender'].get() == selected_gender)

            if event_match and division_match and gender_match:
                row_widgets['frame'].pack(fill='x', pady=2, anchor='w')
            else:
                row_widgets['frame'].pack_forget()

    def _clear_filters(self):
        self.event_filter_combo.set("종목 필터")
        self.division_filter_combo.set("참가부 필터")
        self.gender_filter_combo.set("성별 필터")
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

        event_var = tk.StringVar()
        event_entry = tk.Entry(row_frame, width=18, textvariable=event_var)
        event_entry.pack(side="left", padx=2)
        event_var.trace_add("write", lambda name, index, mode, entry=event_entry: self._update_event_color(entry))

        division_var = tk.StringVar()
        division_entry = tk.Entry(row_frame, width=18, textvariable=division_var)
        division_entry.pack(side="left", padx=2)
        division_var.trace_add("write", lambda name, index, mode, entry=division_entry: self._update_division_color(entry))

        category_entry = tk.Entry(row_frame, width=15)
        category_entry.pack(side="left", padx=2)

        gender_entry = tk.Entry(row_frame, width=10)
        gender_entry.pack(side="left", padx=2)

        count_var = tk.StringVar()
        count_entry = tk.Entry(row_frame, width=10, textvariable=count_var)
        count_entry.pack(side="left", padx=2)
        count_var.trace_add("write", lambda name, index, mode, event_entry=event_entry, count_entry=count_entry: self._update_headcount_color(event_entry, count_entry))
        
        # Bind Tab key to move focus to the next game count entry
        count_entry.bind("<Tab>", self.focus_next_game_count)
        # Bind Shift-Tab for reverse navigation
        count_entry.bind("<Shift-Tab>", self.focus_prev_game_count)

        delete_button = tk.Button(row_frame, text="-", command=lambda: self.remove_input_row(row_widgets))
        delete_button.pack(side="left", padx=2)

        row_widgets = {
            'frame': row_frame, 'check_var': check_var, 'event': event_entry, 
            'division': division_entry, 'class': category_entry, 'gender': gender_entry, 'count': count_entry
        }
        self.input_rows.append(row_widgets)

        if data:
            event_entry.insert(0, data.get("종목", ""))
            division_entry.insert(0, data.get("참가부", ""))
            category_entry.insert(0, data.get("세부부별", ""))
            gender_entry.insert(0, data.get("성별", ""))
            count_entry.insert(0, str(data.get("인원수", "")))
        self._apply_filters()

    def _update_headcount_color(self, event_entry, count_entry):
        event_text = event_entry.get()
        count_text = count_entry.get()

        try:
            original_headcount = int(count_text)
        except ValueError:
            count_entry.config(bg="white", fg="black")
            return

        calculated_headcount = original_headcount

        if event_text == "개인전":
            pass # calculated_headcount remains original_headcount
        elif event_text == "복식전" or event_text == "복식전(자유품새)":
            calculated_headcount = original_headcount / 2
        elif event_text == "단체전":
            calculated_headcount = original_headcount / 3
        elif event_text == "단체전(자유품새)":
            calculated_headcount = original_headcount / 5

        if calculated_headcount < 4 and original_headcount > 0:
            count_entry.config(bg="red", fg="white")
        else:
            count_entry.config(bg="white", fg="black")

    def _update_event_color(self, entry):
        event_text = entry.get()
        color_map = {
            "개인전": "#E0FFFF", # Light Cyan
            "복식전": "#F08080", # Light Coral
            "단체전": "#FAFAD2", # Light Goldenrod Yellow
            "개인전(자유품새)": "#98FB98", # Pale Green
            "복식전(자유품새)": "#FFA07A", # Light Salmon
            "단체전(자유품새)": "#B0C4DE"  # Light Steel Blue
        }
        entry.config(bg=color_map.get(event_text, "white"))
        # Also update headcount color when event changes
        for row_widgets in self.input_rows:
            if row_widgets['event'] == entry:
                self._update_headcount_color(row_widgets['event'], row_widgets['count'])
                break

    def _update_division_color(self, entry):
        division_text = entry.get()
        color_map = {
            "초등부": "#ADD8E6", # Light Blue
            "중등부": "#90EE90", # Light Green
            "고등부": "#FFFFE0", # Light Yellow
            "대학부": "#FFDAB9", # Peach Puff
            "일반부": "#E6E6FA"  # Lavender
        }
        entry.config(bg=color_map.get(division_text, "white"))

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
        download_template_file(TEMPLATE_PATH, "품새_경기시간_계산_양식.xlsx", [("Excel files", "*.xlsx"), ("All files", "*.* ")])

    def open_poomsae_settings(self):
        settings_window = tk.Toplevel(self)
        settings_window.title("품새 옵션")
        settings_window.geometry("500x600")
        settings_window.transient(self)
        settings_window.grab_set()

        self.entries = {}

        canvas = tk.Canvas(settings_window)
        scrollbar = tk.Scrollbar(settings_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        main_frame = scrollable_frame

        def create_time_entry(parent, key, label_text):
            frame = tk.Frame(parent)
            frame.pack(fill="x", pady=2)
            tk.Label(frame, text=label_text, width=12, anchor="w").pack(side="left")
            var = tk.StringVar()
            entry = tk.Entry(frame, width=10, textvariable=var)
            entry.pack(side="left", padx=5)
            tk.Label(frame, text="초").pack(side="left")
            min_label = tk.Label(frame, text="", width=10, fg="gray")
            min_label.pack(side="left", padx=5)
            def update_minutes(*args):
                try:
                    seconds = float(var.get())
                    min_label.config(text=f"({seconds / 60:.1f}분)")
                except (ValueError, tk.TclError):
                    min_label.config(text="")
            var.trace_add("write", update_minutes)
            self.entries[key] = var
            return var

        self.entries['individual'] = {}
        self.entries['team'] = {}
        self.entries['freestyle'] = {}

        gongin_frame = tk.LabelFrame(main_frame, text="공인 시간 설정", padx=10, pady=10)
        gongin_frame.pack(fill="x", padx=15, pady=10)
        individual_frame = tk.LabelFrame(gongin_frame, text="개인전", padx=5, pady=5)
        individual_frame.pack(fill="x", padx=5, pady=5)
        for cat in ["초등부", "중등부", "고등부", "대학부", "일반부"]:
            create_time_entry(individual_frame, f"individual_{cat}", cat)

        team_frame = tk.LabelFrame(gongin_frame, text="복식/단체전", padx=5, pady=5)
        team_frame.pack(fill="x", padx=5, pady=5)
        for cat in ["초등부", "중등부", "고등부", "대학부", "일반부"]:
            create_time_entry(team_frame, f"team_{cat}", cat)

        jayu_frame = tk.LabelFrame(main_frame, text="자유 시간 설정", padx=10, pady=10)
        jayu_frame.pack(fill="x", padx=15, pady=10)
        for cat in DEFAULT_SETTINGS['freestyle']:
            create_time_entry(jayu_frame, f"freestyle_{cat}", cat)

        self.load_settings()

        button_frame = tk.Frame(settings_window)
        button_frame.pack(fill='x', side='bottom', pady=10, padx=15)
        tk.Button(button_frame, text="저장", command=lambda: self.save_settings(settings_window)).pack(side="right", padx=(5,0))
        tk.Button(button_frame, text="기본값", command=lambda: self.restore_defaults(settings_window)).pack(side="right")
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _get_settings_from_ui(self):
        settings = {
            'individual': {cat: self.entries[f'individual_{cat}'].get() for cat in ["초등부", "중등부", "고등부", "대학부", "일반부"]},
            'team': {cat: self.entries[f'team_{cat}'].get() for cat in ["초등부", "중등부", "고등부", "대학부", "일반부"]},
            'freestyle': {cat: self.entries[f'freestyle_{cat}'].get() for cat in DEFAULT_SETTINGS['freestyle']}
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

        for cat in ["초등부", "중등부", "고등부", "대학부", "일반부"]:
            self.entries[f'individual_{cat}'].set(settings_to_load.get('individual', {}).get(cat, ''))
        for cat in ["초등부", "중등부", "고등부", "대학부", "일반부"]:
            self.entries[f'team_{cat}'].set(settings_to_load.get('team', {}).get(cat, ''))
        for cat in DEFAULT_SETTINGS['freestyle']:
            self.entries[f'freestyle_{cat}'].set(settings_to_load.get('freestyle', {}).get(cat, ''))

    def on_close(self):
        self.master.deiconify()
        self.destroy()