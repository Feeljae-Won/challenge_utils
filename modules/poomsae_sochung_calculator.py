import tkinter as tk
from tkinter import ttk, messagebox, font
import math
from decimal import Decimal, getcontext, ROUND_HALF_UP

__version__ = "1.0.0"
__build_date__ = "2025년 7월 6일"

class PoomsaeSochungCalculator(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title(f"품새 소청 계산기 v{__version__} (빌드: {__build_date__})")
        self.geometry("1400x850") # 초기 창 크기 조정
        self.master = master

        # Set precision for Decimal calculations
        getcontext().prec = 10 # Sufficient precision for our needs

        self.judge_count_var = tk.IntVar(value=5) # 기본 5심제
        self.score_system_var = tk.StringVar(value="우리스포츠") # 기본 점수 방식

        # Define styles for colored frames
        self.style = ttk.Style()
        self.style.configure("Blue.TFrame", background="#87CEEB") # SkyBlue
        self.style.configure("Red.TFrame", background="#CD5C5C") # IndianRed

        # Font for LabelFrame titles (1품새, 2품새, 청 선수, 홍 선수)
        self.style.configure("TLabelframe", labelanchor="n") # Center the LabelFrame title
        self.style.configure("TLabelframe.Label", font=("Helvetica", 13, "bold"))

        self.score_entries = {} # 점수 입력 필드를 저장할 딕셔너리
        self.combined_score_displays = {"cheong": {}, "hong": {}} # 합산 점수 표시 레이블 저장

        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        self.destroy()

    def _truncate_float_value(self, number, decimals):
        if not isinstance(number, (int, float, Decimal)):
            return number # Return as is if not a number
        if not isinstance(decimals, int) or decimals < 0:
            raise ValueError("Decimals must be a non-negative integer")

        # Convert to Decimal for precise truncation
        d_number = Decimal(str(number)) # Convert float to string first to avoid binary representation issues
        factor = Decimal(10 ** decimals)
        return (d_number * factor).to_integral_value(rounding='ROUND_DOWN') / factor

    def _apply_taekwondo_soft_sum_rounding(self, number):
        d_number = Decimal(str(number))
        # Check if the number has 1 decimal place (e.g., 5.0, 5.1)
        if d_number == d_number.quantize(Decimal('1.0')):
            return self._round_half_up_value(d_number, 2) # Round to 2 decimal places
        # Check if the number has 2 decimal places (e.g., 5.00, 5.12)
        elif d_number == d_number.quantize(Decimal('1.00')):
            return self._round_half_up_value(d_number, 3) # Round to 3 decimal places
        else:
            return d_number # No specific rounding for other cases

    def _round_half_up_value(self, number, decimals):
        if not isinstance(number, (int, float, Decimal)):
            return number
        if not isinstance(decimals, int) or decimals < 0:
            raise ValueError("Decimals must be a non-negative integer")
        
        # Create a quantize pattern for rounding
        quantize_pattern = Decimal('1.' + '0' * decimals) if decimals > 0 else Decimal('1')
        return Decimal(str(number)).quantize(quantize_pattern, rounding=ROUND_HALF_UP)

    def _format_number_display(self, number, decimals):
        if not isinstance(number, (int, float, Decimal)):
            return str(number)
        if not isinstance(decimals, int) or decimals < 0:
            raise ValueError("Decimals must be a non-negative integer")

        score_system = self.score_system_var.get()

        if score_system == "우리스포츠":
            # 우리스포츠는 버림
            processed_number = self._truncate_float_value(number, decimals)
        else: # 태권소프트
            # 태권소프트는 반올림
            processed_number = self._round_half_up_value(number, decimals)

        return f"{processed_number:.{decimals}f}"

    def create_widgets(self):
        top_frame = ttk.Frame(self)
        top_frame.pack(pady=10, padx=10, fill="x")

        # 심판 수 선택 프레임
        judge_frame = ttk.LabelFrame(top_frame, text="심판 수 선택")
        judge_frame.pack(side="left", padx=5, fill="x", expand=True)

        ttk.Radiobutton(judge_frame, text="3심제", variable=self.judge_count_var, value=3, command=self.update_judge_inputs).pack(side="left", padx=5)
        ttk.Radiobutton(judge_frame, text="5심제", variable=self.judge_count_var, value=5, command=self.update_judge_inputs).pack(side="left", padx=5)
        ttk.Radiobutton(judge_frame, text="7심제", variable=self.judge_count_var, value=7, command=self.update_judge_inputs).pack(side="left", padx=5)

        # 점수 방식 선택 프레임
        score_system_frame = ttk.LabelFrame(top_frame, text="점수 방식 선택")
        score_system_frame.pack(side="left", padx=5, fill="x", expand=True)

        ttk.Radiobutton(score_system_frame, text="우리스포츠", variable=self.score_system_var, value="우리스포츠", command=self.update_scoring_system_info).pack(side="left", padx=5)
        ttk.Radiobutton(score_system_frame, text="태권소프트", variable=self.score_system_var, value="태권소프트", command=self.update_scoring_system_info).pack(side="left", padx=5)

        

        # 1품새 영역
        poomsae1_container_frame = ttk.LabelFrame(self, text="1품새")
        poomsae1_container_frame.pack(pady=10, padx=10, fill="both", expand=True)
        self.create_poomsae_section(poomsae1_container_frame, "poomsae1")

        # 2품새 영역
        poomsae2_container_frame = ttk.LabelFrame(self, text="2품새")
        poomsae2_container_frame.pack(pady=10, padx=10, fill="both", expand=True)
        self.create_poomsae_section(poomsae2_container_frame, "poomsae2")

        # 합산 결과 프레임
        combined_results_frame = ttk.LabelFrame(self, text="합산 결과")
        combined_results_frame.pack(pady=10, padx=10, fill="both", expand=True)

        combined_results_frame.grid_columnconfigure(0, weight=1)
        combined_results_frame.grid_columnconfigure(1, weight=1)

        # 청 선수 합산
        cheong_combined_frame = ttk.LabelFrame(combined_results_frame, text="청 선수 합산")
        cheong_combined_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.create_combined_section(cheong_combined_frame, "cheong")

        # 홍 선수 합산
        hong_combined_frame = ttk.LabelFrame(combined_results_frame, text="홍 선수 합산")
        hong_combined_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        self.create_combined_section(hong_combined_frame, "hong")

        # 승패 표시 레이블
        self.win_loss_label = ttk.Label(combined_results_frame, text="승패: -", font=("Helvetica", 16, "bold"))
        self.win_loss_label.grid(row=1, column=0, columnspan=2, pady=10)

        # 하단 전체 컨테이너 프레임
        bottom_container_frame = ttk.Frame(self)
        bottom_container_frame.pack(pady=10, padx=10, fill="x")

        # 설명 필드셋
        description_frame = ttk.LabelFrame(bottom_container_frame, text="설명")
        description_frame.pack(side="left", fill="both", expand=True, padx=(0, 5)) # 50% 너비

        self.description_label = ttk.Label(description_frame, text="", justify="left") # wraplength 조정
        self.description_label.pack(padx=10, pady=10, fill="both", expand=True)

        # Style for the calculate button
        self.style.configure("Red.TButton", foreground="red", background="red", font=("Helvetica", 15, "bold"))

        # 계산 버튼
        self.calculate_button = ttk.Button(bottom_container_frame, text="", command=self.calculate_all_scores, style="Red.TButton")
        self.calculate_button.pack(side="right", fill="both", expand=True, padx=(5, 0)) # 50% 너비

        self.update_judge_inputs() # 초기 심판 수에 맞춰 입력 필드 생성
        self.update_scoring_system_info() # 초기 버튼 및 설명 텍스트 설정

        # Footer
        footer_font = font.Font(family="Helvetica", size=9)
        footer_label = ttk.Label(self, text="Copyright (c) FEELJAE-WON. All rights reserved.", font=footer_font, foreground="gray")
        footer_label.pack(side=tk.BOTTOM, pady=5)

    def create_poomsae_section(self, parent_frame, poomsae_key):
        # 청/홍 선수 프레임
        cheong_frame = ttk.LabelFrame(parent_frame, text="청 선수")
        cheong_frame.pack(side="left", expand=True, fill="both", padx=5, pady=5)
        self.create_competitor_section(cheong_frame, poomsae_key, "cheong")

        hong_frame = ttk.LabelFrame(parent_frame, text="홍 선수")
        hong_frame.pack(side="right", expand=True, fill="both", padx=5, pady=5)
        self.create_competitor_section(hong_frame, poomsae_key, "hong")

    def create_combined_section(self, parent_frame, competitor_key):
        parent_frame.grid_columnconfigure(0, weight=1) # Item name (e.g., 표현성 합계)
        parent_frame.grid_columnconfigure(1, weight=1) # Item value
        parent_frame.grid_columnconfigure(2, weight=1) # Combined Total Score
        parent_frame.grid_columnconfigure(3, weight=1) # Combined Overall Score

        # Header Row (Row 0)
        # Column 0 will be empty for the header row, as item names are in subsequent rows.
        ttk.Label(parent_frame, text="합계 총점", anchor="center").grid(row=0, column=2, padx=2, pady=2, sticky="nsew")
        ttk.Label(parent_frame, text="합계 평점", anchor="center").grid(row=0, column=3, padx=2, pady=2, sticky="nsew")

        # Row 1: Combined Expression Average
        ttk.Label(parent_frame, text="표현성 합계", anchor="center").grid(row=1, column=0, padx=2, pady=2, sticky="nsew")
        self.combined_score_displays[competitor_key]["combined_expression_avg"] = ttk.Label(parent_frame, text="-", anchor="center")
        self.combined_score_displays[competitor_key]["combined_expression_avg"].grid(row=1, column=1, padx=2, pady=2, sticky="nsew")

        # Row 2: Combined Accuracy Average
        ttk.Label(parent_frame, text="정확성 합계", anchor="center").grid(row=2, column=0, padx=2, pady=2, sticky="nsew")
        self.combined_score_displays[competitor_key]["combined_accuracy_avg"] = ttk.Label(parent_frame, text="-", anchor="center")
        self.combined_score_displays[competitor_key]["combined_accuracy_avg"].grid(row=2, column=1, padx=2, pady=2, sticky="nsew")

        # Combined Total Score (spanning 2 rows) - now in column 2
        self.combined_score_displays[competitor_key]["combined_total_score"] = ttk.Label(parent_frame, text="-", anchor="center", font=("Helvetica", 10, "bold"))
        self.combined_score_displays[competitor_key]["combined_total_score"].grid(row=1, column=2, rowspan=2, padx=2, pady=2, sticky="nsew")

        # Combined Overall Score (spanning 2 rows) - now in column 3
        self.combined_score_displays[competitor_key]["combined_overall_score"] = ttk.Label(parent_frame, text="-", anchor="center", font=("Helvetica", 10, "bold"))
        self.combined_score_displays[competitor_key]["combined_overall_score"].grid(row=1, column=3, rowspan=2, padx=2, pady=2, sticky="nsew")


    def create_competitor_section(self, parent_frame, poomsae_key, competitor_key):
        # Use grid for better layout control
        if competitor_key == "cheong":
            competitor_inner_frame = ttk.Frame(parent_frame, style="Blue.TFrame")
        else:
            competitor_inner_frame = ttk.Frame(parent_frame, style="Red.TFrame")
        
        competitor_inner_frame.pack(expand=True, fill="both")
        competitor_inner_frame.grid_columnconfigure(0, weight=1) # Item label column
        for i in range(1, 8): # Judge columns
            competitor_inner_frame.grid_columnconfigure(i, weight=1)
        competitor_inner_frame.grid_columnconfigure(8, weight=1) # 항목평점 column
        competitor_inner_frame.grid_columnconfigure(9, weight=1) # 총점 column
        competitor_inner_frame.grid_columnconfigure(10, weight=1) # 평점 column


        # Header Row
        # |항목|J1|J2|J3|J4|J5|J6|J7| 항목평점 | 총점 | 평점 |
        ttk.Label(competitor_inner_frame, text="항목", anchor="center").grid(row=0, column=0, padx=2, pady=2, sticky="nsew")
        for i in range(7):
            ttk.Label(competitor_inner_frame, text=f"J{i+1}", anchor="center").grid(row=0, column=i+1, padx=2, pady=2, sticky="nsew")
        ttk.Label(competitor_inner_frame, text="항목평점", anchor="center").grid(row=0, column=8, padx=2, pady=2, sticky="nsew")
        ttk.Label(competitor_inner_frame, text="총점", anchor="center").grid(row=0, column=9, padx=2, pady=2, sticky="nsew")
        ttk.Label(competitor_inner_frame, text="평점", anchor="center").grid(row=0, column=10, padx=2, pady=2, sticky="nsew")

        self.score_entries[poomsae_key] = self.score_entries.get(poomsae_key, {})
        self.score_entries[poomsae_key][competitor_key] = {}

        score_types = ["표현성1 (2.0)", "표현성2 (2.0)", "표현성3 (2.0)", "정확성 (4.0)"]
        row_start = 1 # Start from row 1 after header

        # Score Input Rows
        for idx, score_type in enumerate(score_types):
            current_row = row_start + idx
            ttk.Label(competitor_inner_frame, text=score_type.split(' ')[0], anchor="w").grid(row=current_row, column=0, padx=2, pady=2, sticky="nsew")
            
            self.score_entries[poomsae_key][competitor_key][score_type] = []
            for i in range(7): # Judge Entry fields
                entry = ttk.Entry(competitor_inner_frame, width=5) # Smaller width for entries
                entry.grid(row=current_row, column=i+1, padx=2, pady=2, sticky="nsew")
                entry.bind("<KeyRelease>", self._on_score_entry_key_release)
                self.score_entries[poomsae_key][competitor_key][score_type].append(entry)
            
            # 항목평점 Label
            if "표현성" in score_type:
                if score_type == "표현성1 (2.0)": # Only create for the first expression type
                    avg_label = ttk.Label(competitor_inner_frame, text="-", anchor="center")
                    avg_label.grid(row=current_row, column=8, rowspan=3, padx=2, pady=2, sticky="nsew") # Span 3 rows
                    self.score_entries[poomsae_key][competitor_key]["expression_item_avg_sum_label"] = avg_label # New key for the sum
                # For Expression2 and Expression3, no separate label is needed here
            else: # For Accuracy
                avg_label = ttk.Label(competitor_inner_frame, text="-", anchor="center")
                avg_label.grid(row=current_row, column=8, padx=2, pady=2, sticky="nsew")
                self.score_entries[poomsae_key][competitor_key][score_type + "_avg_label"] = avg_label

        # Special Spanning Labels
        # 표현성 전체의 합계 (spanning Expression1, Expression2, Expression3 rows)
        self.score_entries[poomsae_key][competitor_key]["expression_overall_avg_display"] = ttk.Label(competitor_inner_frame, text="표현성 합계: -", anchor="center")
        self.score_entries[poomsae_key][competitor_key]["expression_overall_avg_display"].grid(row=row_start, column=9, rowspan=3, padx=2, pady=2, sticky="nsew") # Starts at Expression1 row, spans 3 rows

        # 총점 (spanning all 4 score rows)
        self.score_entries[poomsae_key][competitor_key]["unweighted_total_score_display"] = ttk.Label(competitor_inner_frame, text="-", anchor="center", font=("Helvetica", 10, "bold"))
        self.score_entries[poomsae_key][competitor_key]["unweighted_total_score_display"].grid(row=row_start, column=9, rowspan=4, padx=2, pady=2, sticky="nsew") # Starts at Expression1 row, spans 4 rows

        # 평점 (spanning all 4 score rows)
        self.score_entries[poomsae_key][competitor_key]["weighted_overall_score_display"] = ttk.Label(competitor_inner_frame, text="-", anchor="center", font=("Helvetica", 10, "bold"))
        self.score_entries[poomsae_key][competitor_key]["weighted_overall_score_display"].grid(row=row_start, column=10, rowspan=4, padx=2, pady=2, sticky="nsew") # Starts at Expression1 row, spans 4 rows

        # Store references to these for updating
        self.score_entries[poomsae_key][competitor_key]["accuracy_avg_display"] = self.score_entries[poomsae_key][competitor_key]["정확성 (4.0)_avg_label"]
        self.score_entries[poomsae_key][competitor_key]["expression_avg_display"] = self.score_entries[poomsae_key][competitor_key]["expression_overall_avg_display"]
        self.score_entries[poomsae_key][competitor_key]["total_score_display"] = self.score_entries[poomsae_key][competitor_key]["unweighted_total_score_display"]
        self.score_entries[poomsae_key][competitor_key]["overall_score_display"] = self.score_entries[poomsae_key][competitor_key]["weighted_overall_score_display"]

    def _on_score_entry_key_release(self, event):
        entry = event.widget
        current_text = entry.get()

        # Remove any non-digit characters except for a single decimal point
        cleaned_text = "".join(filter(lambda x: x.isdigit() or x == ".", current_text))
        if cleaned_text.count(".") > 1:
            # If more than one decimal point, keep only the first one
            parts = cleaned_text.split(".", 1)
            cleaned_text = parts[0] + "." + parts[1].replace(".", "")

        if cleaned_text != current_text:
            entry.delete(0, tk.END)
            entry.insert(0, cleaned_text)
            current_text = cleaned_text

        if not current_text:
            return

        try:
            # Try to convert to float to check if it's a valid number
            value = float(current_text)

            # If the value is 10 or greater and doesn't contain a decimal point
            if value >= 10 and "." not in current_text:
                # Insert a decimal point before the last digit
                new_text = current_text[:-1] + "." + current_text[-1]
                entry.delete(0, tk.END)
                entry.insert(0, new_text)
        except ValueError:
            # Not a valid number, do nothing or show an error
            pass


    def update_judge_inputs(self):
        current_judge_count = self.judge_count_var.get()
        # Reset all entry backgrounds to white
        for poomsae_key in ["poomsae1", "poomsae2"]:
            for competitor_key in ["cheong", "hong"]:
                for score_type in ["정확성 (4.0)", "표현성1 (2.0)", "표현성2 (2.0)", "표현성3 (2.0)"]:
                    for entry in self.score_entries[poomsae_key][competitor_key][score_type]:
                        entry.config(background="white")

        for poomsae_key in ["poomsae1", "poomsae2"]:
            for competitor_key in ["cheong", "hong"]:
                for score_type in ["정확성 (4.0)", "표현성1 (2.0)", "표현성2 (2.0)", "표현성3 (2.0)"]:
                    for i, entry in enumerate(self.score_entries[poomsae_key][competitor_key][score_type]):
                        if i < current_judge_count:
                            entry.config(state="normal")
                        else:
                            entry.config(state="disabled")
                            entry.delete(0, tk.END) # 비활성화 시 내용 삭제
        self.calculate_all_scores() # Recalculate scores after judge count update

    def update_scoring_system_info(self):
        score_system = self.score_system_var.get()
        if score_system == "우리스포츠":
            self.calculate_button.config(text="우리스포츠 방식 점수 계산")
            self.description_label.config(text="* 우리스포츠 방식 점수 계산\n   - 우리스포츠는 항목평점 및 평점을 계산할 때 소수점 3자리 까지 표현되며 소수점 4자리에서 내림 합니다. \n * 동점 처리 \n   - 1) 표현력 > 2) 정확성 > 3) 총점")
        else: # 태권소프트
            self.calculate_button.config(text="태권소프트 방식 점수 계산")
            self.description_label.config(text="* 태권소프트 방식 점수 계산\n   - 태권소프트는 항목평점 및 평점 계산 할 때 아래 식에 따릅니다.\n   - 각 품새 평점 : 소수점 3자리에서 반올림.\n   - 총 평점 : 소수점 4자리에서 반올림.\n   - 총점 : 소수점 3자리에서 반올림.  \n * 동점 처리 \n   - 1) 표현력 > 2) 정확성 > 3) 총점")
        self.calculate_all_scores()

    def calculate_all_scores(self):
        # Reset all entry backgrounds to white before calculation
        for poomsae_key in ["poomsae1", "poomsae2"]:
            for competitor_key in ["cheong", "hong"]:
                for score_type in ["정확성 (4.0)", "표현성1 (2.0)", "표현성2 (2.0)", "표현성3 (2.0)"]:
                    for entry in self.score_entries[poomsae_key][competitor_key][score_type]:
                        entry.config(background="white")

        poomsae_results = {}
        valid_poomsae_counts = {"cheong": 0, "hong": 0}

        for poomsae_key in ["poomsae1", "poomsae2"]:
            poomsae_results[poomsae_key] = {}
            for competitor_key in ["cheong", "hong"]:
                results = self.calculate_competitor_scores(poomsae_key, competitor_key)
                poomsae_results[poomsae_key][competitor_key] = results
                self.display_competitor_scores(poomsae_key, competitor_key, results)
                if results.get("sum_of_item_averages", Decimal('0.0')) > Decimal('0.0'):
                    valid_poomsae_counts[competitor_key] += 1
        
        # 합산 점수 계산 및 표시
        cheong_combined_expression_avg = Decimal('0.0')
        cheong_combined_accuracy_avg = Decimal('0.0')
        cheong_combined_unweighted_total_score = Decimal('0.0')
        cheong_combined_final_score = Decimal('0.0')

        hong_combined_expression_avg = Decimal('0.0')
        hong_combined_accuracy_avg = Decimal('0.0')
        hong_combined_unweighted_total_score = Decimal('0.0')
        hong_combined_final_score = Decimal('0.0')

        for poomsae_key in ["poomsae1", "poomsae2"]:
            cheong_results = poomsae_results[poomsae_key]["cheong"]
            hong_results = poomsae_results[poomsae_key]["hong"]

            if cheong_results.get("sum_of_item_averages") is not None:
                cheong_combined_expression_avg += cheong_results["sum_of_expression_item_averages"]
                cheong_combined_accuracy_avg += cheong_results["accuracy_avg"]
                cheong_combined_unweighted_total_score += cheong_results["raw_judge_scores_sum"]
                cheong_combined_final_score += cheong_results["sum_of_item_averages"]
            
            if hong_results.get("sum_of_item_averages") is not None:
                hong_combined_expression_avg += hong_results["sum_of_expression_item_averages"]
                hong_combined_accuracy_avg += hong_results["accuracy_avg"]
                hong_combined_unweighted_total_score += hong_results["raw_judge_scores_sum"]
                hong_combined_final_score += hong_results["sum_of_item_averages"]

        # Average the scores only if more than one poomsae is valid
        cheong_divisor = valid_poomsae_counts["cheong"]
        if cheong_divisor > 0: # Check for division by zero
            cheong_combined_expression_avg /= cheong_divisor
            cheong_combined_accuracy_avg /= cheong_divisor
            cheong_combined_unweighted_total_score /= cheong_divisor
            cheong_combined_final_score /= cheong_divisor

        hong_divisor = valid_poomsae_counts["hong"]
        if hong_divisor > 0: # Check for division by zero
            hong_combined_expression_avg /= hong_divisor
            hong_combined_accuracy_avg /= hong_divisor
            hong_combined_unweighted_total_score /= hong_divisor
            hong_combined_final_score /= hong_divisor

        # Display combined scores for Cheong
        self.combined_score_displays["cheong"]["combined_expression_avg"].config(text=self._format_number_display(cheong_combined_expression_avg, 3))
        self.combined_score_displays["cheong"]["combined_accuracy_avg"].config(text=self._format_number_display(cheong_combined_accuracy_avg, 3))
        
        score_system = self.score_system_var.get()
        if score_system == "우리스포츠":
            self.combined_score_displays["cheong"]["combined_total_score"].config(text=self._format_number_display(cheong_combined_unweighted_total_score, 1))
            self.combined_score_displays["cheong"]["combined_overall_score"].config(text=self._format_number_display(self._truncate_float_value(cheong_combined_final_score, 3), 3))
        else: # 태권소프트
            self.combined_score_displays["cheong"]["combined_total_score"].config(text=self._format_number_display(self._round_half_up_value(cheong_combined_unweighted_total_score, 2), 2))
            self.combined_score_displays["cheong"]["combined_overall_score"].config(text=self._format_number_display(self._round_half_up_value(cheong_combined_final_score, 3), 3))

        # Display combined scores for Hong
        self.combined_score_displays["hong"]["combined_expression_avg"].config(text=self._format_number_display(hong_combined_expression_avg, 3))
        self.combined_score_displays["hong"]["combined_accuracy_avg"].config(text=self._format_number_display(hong_combined_accuracy_avg, 3))

        if score_system == "우리스포츠":
            self.combined_score_displays["hong"]["combined_total_score"].config(text=self._format_number_display(hong_combined_unweighted_total_score, 1))
            self.combined_score_displays["hong"]["combined_overall_score"].config(text=self._format_number_display(self._truncate_float_value(hong_combined_final_score, 3), 3))
        else: # 태권소프트
            self.combined_score_displays["hong"]["combined_total_score"].config(text=self._format_number_display(self._round_half_up_value(hong_combined_unweighted_total_score, 2), 2))
            self.combined_score_displays["hong"]["combined_overall_score"].config(text=self._format_number_display(self._round_half_up_value(hong_combined_final_score, 3), 3))

        # Determine Win/Loss based on the final calculated score with tie-breaking rules
        final_cheong_score = self._truncate_float_value(cheong_combined_final_score, 3)
        final_hong_score = self._truncate_float_value(hong_combined_final_score, 3)

        if final_cheong_score > final_hong_score:
            self.win_loss_label.config(text="승패: 청 선수 승!", foreground="navy")
        elif final_hong_score > final_cheong_score:
            self.win_loss_label.config(text="승패: 홍 선수 승!", foreground="red")
        else: # 합계 평점 동점일 경우
            # 표현성 합계 비교
            if cheong_combined_expression_avg > hong_combined_expression_avg:
                self.win_loss_label.config(text="승패: 청 선수 승! (표현성 우위)", foreground="navy")
            elif hong_combined_expression_avg > cheong_combined_expression_avg:
                self.win_loss_label.config(text="승패: 홍 선수 승! (표현성 우위)", foreground="red")
            else: # 표현성 합계도 동점일 경우
                # 정확성 합계 비교
                if cheong_combined_accuracy_avg > hong_combined_accuracy_avg:
                    self.win_loss_label.config(text="승패: 청 선수 승! (정확성 우위)", foreground="navy")
                elif hong_combined_accuracy_avg > cheong_combined_accuracy_avg:
                    self.win_loss_label.config(text="승패: 홍 선수 승! (정확성 우위)", foreground="red")
                else: # 정확성 합계도 동점일 경우
                    # 합계 총점 비교
                    if cheong_combined_unweighted_total_score > hong_combined_unweighted_total_score:
                        self.win_loss_label.config(text="승패: 청 선수 승! (총점 우위)", foreground="navy")
                    elif hong_combined_unweighted_total_score > cheong_combined_unweighted_total_score:
                        self.win_loss_label.config(text="승패: 홍 선수 승! (총점 우위)", foreground="red")
                    else: # 모든 항목 동점일 경우
                        self.win_loss_label.config(text="승패: 무승부", foreground="black")


    def calculate_competitor_scores(self, poomsae_key, competitor_key):
        current_judge_count = self.judge_count_var.get()
        score_system = self.score_system_var.get()
        
        item_averages = {} # To store averages for each score type (정확성, 표현성1, 표현성2, 표현성3)
        expression_item_averages = [] # To store averages for Expression1, Expression2, Expression3
        raw_judge_scores_sum = Decimal('0.0') # New: Sum of all raw judge scores
        all_excluded_entries = [] # To store references to excluded entry widgets

        score_types = ["표현성1 (2.0)", "표현성2 (2.0)", "표현성3 (2.0)", "정확성 (4.0)"]

        for score_type_idx, score_type in enumerate(score_types):
            judge_scores_with_entries = [] # Store (score, entry_widget) tuples
            for i in range(current_judge_count):
                entry = self.score_entries[poomsae_key][competitor_key][score_type][i]
                entry_value = entry.get()
                if not entry_value.strip(): # Check if empty or only whitespace
                    score = Decimal('0.0')
                else:
                    try:
                        score = Decimal(entry_value)
                        if not (Decimal('0') <= score <= Decimal('10')):
                            messagebox.showerror("입력 오류", f"{poomsae_key} {competitor_key} {score_type} 심판{i+1} 점수는 0에서 10 사이여야 합니다.")
                            score = Decimal('0.0') # Use 0 for calculation if invalid range
                    except Exception: # Catch all exceptions for Decimal conversion
                        messagebox.showerror("입력 오류", f"{poomsae_key} {competitor_key} {score_type} 심판{i+1} 점수가 유효한 숫자가 아닙니다.")
                        score = Decimal('0.0') # Use 0 for calculation if invalid type
                judge_scores_with_entries.append((score, entry))
                raw_judge_scores_sum += score # Accumulate raw score
            
            # 최고, 최저점 제외 (5심제, 7심제)
            excluded_entries_for_type = []
            if current_judge_count in [5, 7] and len(judge_scores_with_entries) >= 2:
                # Sort by score to find min/max
                judge_scores_with_entries.sort(key=lambda x: x[0])
                
                # Identify excluded entries
                excluded_entries_for_type.append(judge_scores_with_entries[0][1]) # Lowest score's entry
                excluded_entries_for_type.append(judge_scores_with_entries[-1][1]) # Highest score's entry

                # Keep only included scores for calculation
                judge_scores = [item[0] for item in judge_scores_with_entries[1:-1]]
            else:
                judge_scores = [item[0] for item in judge_scores_with_entries]

            if judge_scores:
                if score_system == "우리스포츠":
                    avg_score = self._truncate_float_value(sum(judge_scores) / Decimal(str(len(judge_scores))), 3)
                else: # 태권소프트
                    avg_score = self._round_half_up_value(sum(judge_scores) / Decimal(str(len(judge_scores))), 2)
            else:
                avg_score = Decimal('0.0') # 점수가 없을 경우 0으로 처리

            item_averages[score_type] = avg_score
            
            if "표현성" in score_type:
                expression_item_averages.append(avg_score)
            else: # For Accuracy
                self.score_entries[poomsae_key][competitor_key][score_type + "_avg_label"].config(text=self._format_number_display(avg_score, 2 if score_system == "태권소프트" else 3))
            
            all_excluded_entries.extend(excluded_entries_for_type)

        # Calculate sum of expression item averages
        sum_of_expression_item_averages = Decimal('0.0')
        if expression_item_averages:
            sum_of_expression_item_averages = sum(expression_item_averages)

        # Calculate expression overall average (average of the three expression scores)
        expression_overall_avg = Decimal('0.0')
        if expression_item_averages:
            if score_system == "우리스포츠":
                expression_overall_avg = self._truncate_float_value(sum_of_expression_item_averages / Decimal(str(len(expression_item_averages))), 3)
            else: # 태권소프트
                expression_overall_avg = self._round_half_up_value(sum_of_expression_item_averages / Decimal(str(len(expression_item_averages))), 2)
        
        # Calculate sum of item averages (sum of expression item averages + accuracy average)
        sum_of_item_averages = sum_of_expression_item_averages + item_averages.get("정확성 (4.0)", Decimal('0.0'))

        return {
            "expression_overall_avg": expression_overall_avg,
            "sum_of_expression_item_averages": sum_of_expression_item_averages, # New return value
            "accuracy_avg": item_averages.get("정확성 (4.0)", Decimal('0.0')),
            "sum_of_item_averages": sum_of_item_averages,
            "raw_judge_scores_sum": raw_judge_scores_sum,
            "excluded_entries": all_excluded_entries # Return excluded entries
        }

    def display_competitor_scores(self, poomsae_key, competitor_key, results):
        score_system = self.score_system_var.get()

        # Update expression item average sum label
        expression_item_avg_sum_label = self.score_entries[poomsae_key][competitor_key].get("expression_item_avg_sum_label")
        if expression_item_avg_sum_label:
            expression_item_avg_sum_label.config(text=self._format_number_display(results["sum_of_expression_item_averages"], 2 if score_system == "태권소프트" else 3))

        # Update accuracy average display
        accuracy_avg_display = self.score_entries[poomsae_key][competitor_key].get("accuracy_avg_display")
        if accuracy_avg_display:
            accuracy_avg_display.config(text=self._format_number_display(results["accuracy_avg"], 2 if score_system == "태권소프트" else 3))

        # Update overall expression average display
        expression_overall_avg_display = self.score_entries[poomsae_key][competitor_key].get("expression_overall_avg_display")
        if expression_overall_avg_display:
            expression_overall_avg_display.config(text=f"표현성 합계: {self._format_number_display(results['expression_overall_avg'], 3)}")

        # Update unweighted total score display
        unweighted_total_score_display = self.score_entries[poomsae_key][competitor_key].get("unweighted_total_score_display")
        if unweighted_total_score_display:
            if score_system == "우리스포츠":
                unweighted_total_score_display.config(text=self._format_number_display(results["raw_judge_scores_sum"], 1))
            else: # 태권소프트
                unweighted_total_score_display.config(text=self._format_number_display(self._round_half_up_value(results["raw_judge_scores_sum"], 1), 1))

        # Update weighted overall score display
        weighted_overall_score_display = self.score_entries[poomsae_key][competitor_key].get("weighted_overall_score_display")
        if weighted_overall_score_display:
            if score_system == "우리스포츠":
                weighted_overall_score_display.config(text=self._format_number_display(self._truncate_float_value(results["sum_of_item_averages"], 3), 3))
            else: # 태권소프트
                weighted_overall_score_display.config(text=self._format_number_display(self._round_half_up_value(results["sum_of_item_averages"], 2), 2))

        # Highlight excluded entries
        for entry in results["excluded_entries"]:
            entry.config(background="yellow")

    