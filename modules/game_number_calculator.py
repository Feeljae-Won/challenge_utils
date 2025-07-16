import tkinter as tk
from tkinter import ttk, font, filedialog
import shutil
import openpyxl
import os
import datetime
from ..common.version import __version__ as app_version
from ..common.version import __build_date__ as app_date
from ..common.constants import GAME_NUMBER_TEMPLATE_PATH
from ..utils.file_operations import download_template_file

class GameNumberCalculator(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.version = app_version
        self.title(f"경기번호 계산기 v{app_version} (빌드: {app_date})")
        self.geometry("1200x700")
        self.last_imported_filename = ""
        self.last_imported_filename = ""

        # 설명 레이블
        description_font = font.Font(family="Helvetica", size=12)
        description_label = tk.Label(self, text="컷오프 계산시 종목에 '자유품새'를 입력하세요.", font=description_font, pady=10)
        description_label.pack()

        # 메인 프레임 (좌우 분할)
        main_container = tk.Frame(self)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 왼쪽: 입력 컨테이너
        input_container = tk.Frame(main_container)
        input_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 버튼 및 헤더 프레임 (입력 컨테이너 내 상단)
        top_controls_frame = tk.Frame(input_container)
        top_controls_frame.pack(fill=tk.X, pady=5)

        # '+' 버튼 (왼쪽 정렬)
        plus_button_font = font.Font(family="Helvetica", size=13, weight="bold")
        add_row_button = tk.Button(top_controls_frame, text="+", bg="#030ac1", fg="white", command=self.add_row, font=plus_button_font)
        add_row_button.pack(side=tk.LEFT, padx=5)

        # 초기화 버튼 (빨간색 바탕, 흰 글씨)
        reset_button = tk.Button(top_controls_frame, text="초기화", bg="red", fg="white", command=self.reset_all)
        reset_button.pack(side=tk.LEFT, padx=5)

        # 오른쪽 정렬 버튼들을 담을 프레임
        right_buttons_frame = tk.Frame(top_controls_frame)
        right_buttons_frame.pack(side=tk.RIGHT)

        # 계산하기 버튼 (3번째, 붉은색)
        calculate_button = tk.Button(right_buttons_frame, text="계산하기", bg="red", fg="white", command=self.calculate_matches)
        calculate_button.pack(side=tk.RIGHT)

        # 엑셀로 가져오기 버튼 (2번째, 녹색)
        import_button = tk.Button(right_buttons_frame, text="엑셀로 가져오기", bg="#4CAF50", fg="white", command=self.import_from_excel)
        import_button.pack(side=tk.RIGHT, padx=5)

        # 엑셀 양식 다운로드 버튼 (1번째, 기본색)
        download_button = tk.Button(right_buttons_frame, text="엑셀 양식 다운로드", command=self.download_template)
        download_button.pack(side=tk.RIGHT, padx=5)

        # 헤더 프레임 (입력 컨테이너 내 상단, 버튼 아래)
        header_frame = tk.Frame(input_container)
        header_frame.pack(fill=tk.X, pady=5)

        header_labels = ["번호", "종목", "부", "체급", "참가인원"]
        for i, label_text in enumerate(header_labels):
            width = 5 if label_text == "번호" else 15
            label = tk.Label(header_frame, text=label_text, width=width)
            label.pack(side=tk.LEFT, padx=5)

        # '+' 버튼을 위한 빈 공간 (헤더와 정렬)
        tk.Label(header_frame, width=5).pack(side=tk.LEFT, padx=5)

        # 스크롤 가능한 입력 행 프레임
        self.canvas = tk.Canvas(input_container)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = ttk.Scrollbar(input_container, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion = self.canvas.bbox("all")))

        self.main_frame = tk.Frame(self.canvas)
        self.canvas.bind('<Configure>', lambda e: self.canvas.itemconfig(self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw", width=e.width), width=e.width))

        self.main_frame.bind('<Configure>', lambda e: self.canvas.configure(scrollregion = self.canvas.bbox("all")))

        # 오른쪽: 결과 프레임
        result_frame = tk.Frame(main_container, width=700)
        result_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10)
        result_frame.pack_propagate(False) #결과 프레임 크기 고정

        result_label = tk.Label(result_frame, text="계산 결과")
        result_label.pack(pady=5)

        # 결과 표시 Treeview
        tree_frame = tk.Frame(result_frame) # Treeview와 스크롤바를 담을 프레임
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.result_tree = ttk.Treeview(tree_frame, columns=("번호", "종목", "부", "체급", "강수", "경기번호", "경기수"), show='tree headings', selectmode='extended')
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.result_tree.yview)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_tree.configure(yscrollcommand=tree_scrollbar.set)

        # 결과 다운로드 버튼
        # 결과 다운로드 버튼
        download_results_button = tk.Button(result_frame, text="결과 다운로드", bg="red", fg="white", font=("Helvetica", 12, "bold"), command=self.export_results_to_excel)
        download_results_button.pack(fill=tk.X, pady=5)

        # 각 열 설정
        self.result_tree.heading("번호", text="번호")
        self.result_tree.column("번호", width=40, anchor='center')
        self.result_tree.column("#0", width=0, stretch=tk.NO) # 트리 열 숨기기
        self.result_tree.heading("종목", text="종목")
        self.result_tree.column("종목", width=120)
        self.result_tree.heading("부", text="부")
        self.result_tree.column("부", width=80)
        self.result_tree.heading("체급", text="체급")
        self.result_tree.column("체급", width=80)
        self.result_tree.heading("강수", text="강수")
        self.result_tree.column("강수", width=70, anchor='center')
        self.result_tree.heading("경기번호", text="경기번호")
        self.result_tree.column("경기번호", width=80, anchor='center')
        self.result_tree.heading("경기수", text="경기수")
        self.result_tree.column("경기수", width=40, anchor='center')

        # 정렬 상태 초기화
        self.sort_state = {col: 0 for col in self.result_tree["columns"]}

        # 헤더 클릭 이벤트 바인딩
        for col in self.result_tree["columns"]:
            self.result_tree.heading(col, command=lambda _col=col: self._sort_column(_col))

        self.rows = []
        for _ in range(10):
            self.add_row()

        # Treeview에 복사 기능 바인딩
        self.result_tree.bind("<Control-c>", self._copy_selected_rows)
        self.result_tree.bind("<Command-c>", self._copy_selected_rows) # For macOS

        # 마우스 휠 스크롤 바인딩
        self.canvas.bind("<MouseWheel>", self._on_mousewheel) # Windows/Linux
        self.canvas.bind("<Button-4>", self._on_mousewheel) # macOS scroll up
        self.result_tree.bind("<Button-5>", self._on_mousewheel) # macOS scroll down

        # Footer
        footer_font = font.Font(family="Helvetica", size=9)
        footer_label = tk.Label(self, text="Copyright (c) FEELJAE-WON. All rights reserved.", font=footer_font, fg="gray")
        footer_label.pack(side=tk.BOTTOM, pady=5)
        self.main_frame.bind("<MouseWheel>", self._on_mousewheel) # Windows/Linux
        self.main_frame.bind("<Button-4>", self._on_mousewheel) # macOS scroll up
        self.main_frame.bind("<Button-5>", self._on_mousewheel) # macOS scroll down

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _copy_selected_rows(self, event=None):
        selected_items = self.result_tree.selection()
        if not selected_items:
            return

        clipboard_content = []
        for item_id in selected_items:
            values = self.result_tree.item(item_id, 'values')
            clipboard_content.append('\t'.join(map(str, values)))
        
        self.clipboard_clear()
        self.clipboard_append('\n'.join(clipboard_content))

    def reset_all(self):
        # 결과 Treeview 초기화
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 입력 필드 초기화
        for row_data in self.rows:
            row_data["frame"].destroy()
        self.rows.clear()

        # 초기 10개 행 추가
        for _ in range(10):
            self.add_row()

    def remove_row(self, row_frame_to_remove):
        for i, row_data in enumerate(self.rows):
            if row_data["frame"] == row_frame_to_remove:
                row_data["frame"].destroy()
                del self.rows[i]
                self.resequence_row_numbers()
                break

    def resequence_row_numbers(self):
        for i, row_data in enumerate(self.rows):
            row_data["number_label"].config(text=str(i + 1))

    def add_row(self):
        row_frame = tk.Frame(self.main_frame)
        row_frame.pack(fill=tk.X, expand=True, pady=5)

        row_number = len(self.rows) + 1
        number_label = tk.Label(row_frame, text=str(row_number), width=5)
        number_label.pack(side=tk.LEFT, padx=5)

        labels = ["종목", "부", "체급", "참가인원"]
        entries = {}

        for i, label_text in enumerate(labels):
            entry = tk.Entry(row_frame, width=15)
            entry.pack(side=tk.LEFT, padx=5)
            entries[label_text] = entry
        
        # '-' 버튼 추가
        remove_button_font = font.Font(family="Helvetica", size=12, weight="bold")
        remove_button = tk.Button(row_frame, text="-", bg="red", fg="white", font=remove_button_font,
                                  command=lambda r_frame=row_frame: self.remove_row(r_frame))
        remove_button.pack(side=tk.LEFT, padx=5)

        self.rows.append({"frame": row_frame, "entries": entries, "number_label": number_label, "remove_button": remove_button})

    def import_from_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        self.last_imported_filename = os.path.splitext(os.path.basename(file_path))[0]

        # Clear existing rows
        for row_data in self.rows:
            row_data["frame"].destroy()
        self.rows.clear()

        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        for r_idx, row in enumerate(sheet.iter_rows(values_only=True)):
            if r_idx == 0: # Skip header row
                continue
            self.add_row_with_data(row)

    def download_template(self):
        download_template_file(GAME_NUMBER_TEMPLATE_PATH, "경기번호_계산기_양식.xlsx", [("Excel files", "*.xlsx")])

    def calculate_matches(self):
        # 기존 결과 삭제
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 정렬 상태 초기화 및 헤더 화살표 제거
        for col in self.result_tree["columns"]:
            self.sort_state[col] = 0
            current_text = self.result_tree.heading(col, "text")
            self.result_tree.heading(col, text=current_text.replace(" ▲", "").replace(" ▼", ""))

        game_number_counter = 1
        row_index = 1 # 결과 테이블의 행 번호
        prev_event, prev_division, prev_weight_class = None, None, None

        for i, row_data in enumerate(self.rows):
            entries = row_data["entries"]
            try:
                participants_str = entries["참가인원"].get()
                if not participants_str.strip():
                    continue

                participants = int(participants_str)
                if participants < 2:
                    continue

                event = entries["종목"].get() or ""
                division = entries["부"].get() or ""
                weight_class = entries["체급"].get() or ""

                if event == "자유품새":
                    if participants <= 11:
                        # Case 1: Participants <= 11
                        self.result_tree.insert("", "end", values=(row_index, event, division, weight_class, "결선", f"1~{participants}", participants))
                        row_index += 1
                    elif 12 <= participants <= 21:
                        # Case 2: 12 <= Participants <= 21
                        # Divide into 2 groups, with the first group being larger if uneven
                        group1_size = (participants + 1) // 2
                        group2_size = participants - group1_size

                        # 본선-1조
                        self.result_tree.insert("", "end", values=(row_index, event, division, weight_class, "본선-1조", f"1~{group1_size}", group1_size))
                        row_index += 1
                        # 본선-2조
                        self.result_tree.insert("", "end", values=(row_index, event, division, weight_class, "본선-2조", f"1~{group2_size}", group2_size))
                        row_index += 1

                        # 결선
                        self.result_tree.insert("", "end", values=(row_index, event, division, weight_class, "결선", "1~8", 8))
                        row_index += 1
                    elif participants >= 22:
                        # Case 3: Participants >= 22
                        # Preliminary (예선)
                        num_prelim_groups = 2 # Start with 2 groups
                        while participants / num_prelim_groups > 11.5 and num_prelim_groups % 2 == 0:
                            num_prelim_groups += 2

                        # Ensure num_prelim_groups is at least 2 and even
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

                        # 예선 결과 출력
                        for g_idx, size in enumerate(prelim_group_sizes):
                            self.result_tree.insert("", "end", values=(row_index, event, division, weight_class, f"예선-{g_idx+1}조", f"1~{size}", size))
                            row_index += 1

                        # 본선 진출 인원 계산 (각 예선 조에서 50% 진출)
                        total_main_round_participants = 0
                        for size in prelim_group_sizes:
                            total_main_round_participants += (size + 1) // 2 # ceil(size / 2) for 50% advancement

                        # 본선 (Main Round)
                        if total_main_round_participants > 0:
                            num_main_groups = 2 # Start with 2 groups
                            while total_main_round_participants / num_main_groups > 11.5 and num_main_groups % 2 == 0:
                                num_main_groups += 2

                            # Ensure num_main_groups is at least 2 and even
                            if num_main_groups < 2:
                                num_main_groups = 2
                            if num_main_groups % 2 != 0:
                                num_main_groups += 1

                            base_main_group_size = total_main_round_participants // num_main_groups
                            remainder_main = total_main_round_participants % num_main_groups

                            main_group_sizes = []
                            for g in range(num_main_groups):
                                size = base_main_group_size
                                if g < remainder_main:
                                    size += 1
                                main_group_sizes.append(size)

                            # 본선 결과 출력
                            for g_idx, size in enumerate(main_group_sizes):
                                self.result_tree.insert("", "end", values=(row_index, event, division, weight_class, f"본선-{g_idx+1}조", f"1~{size}", size))
                                row_index += 1

                            # 결선 (Final Round) - based on number of main round groups
                            # 예선과 본선을 거쳤을 경우 결선은 무조건 1~8
                            self.result_tree.insert("", "end", values=(row_index, event, division, weight_class, "결선", "1~8", 8))
                            row_index += 1
                    
                    prev_event, prev_division, prev_weight_class = None, None, None

                else:
                    # Existing logic for other events
                    current_category = (event, division, weight_class)
                    previous_category = (prev_event, prev_division, prev_weight_class)

                    if current_category != previous_category:
                        game_number_counter = 1

                    game_number_counter, row_index = self._calculate_standard_matches(participants, event, division, weight_class, game_number_counter, row_index)
                    
                    prev_event, prev_division, prev_weight_class = event, division, weight_class

            except ValueError:
                continue

    def _calculate_standard_matches(self, participants, event, division, weight_class, game_number_counter, row_index):
        total_slots = 1
        while total_slots < participants:
            total_slots *= 2

        byes = total_slots - participants
        first_round_matches = participants - byes

        # 예선전 (첫 라운드)
        if first_round_matches > 0:
            round_name = f"{total_slots}"
            num_matches = first_round_matches // 2
            start_game = game_number_counter
            end_game = game_number_counter + num_matches - 1
            game_numbers_display = f"{start_game}~{end_game}" if num_matches > 0 else "-"
            self.result_tree.insert("", "end", values=(row_index, event, division, weight_class, round_name, game_numbers_display, num_matches))
            game_number_counter = end_game + 1
            row_index += 1

        # 본선 (다음 라운드부터 결승까지)
        current_participants = (first_round_matches // 2) + byes
        while current_participants > 1:
            round_matches = current_participants // 2
            round_name = f"{current_participants}" if current_participants > 2 else "2"
            if current_participants == 4:
                round_name = "4"
            elif current_participants == 8:
                round_name = "8"

            start_game = game_number_counter
            end_game = game_number_counter + round_matches - 1
            game_numbers_display = f"{start_game}~{end_game}" if round_matches > 0 else "-"
            self.result_tree.insert("", "end", values=(row_index, event, division, weight_class, round_name, game_numbers_display, round_matches))
            game_number_counter = end_game + 1
            current_participants //= 2
            row_index += 1
        return game_number_counter, row_index

    def _sort_column(self, col):
        # 현재 열의 정렬 상태 업데이트
        current_state = self.sort_state[col]
        
        # 모든 헤더에서 화살표 제거
        for c in self.result_tree["columns"]:
            current_text = self.result_tree.heading(c, "text")
            self.result_tree.heading(c, text=current_text.replace(" ▲", "").replace(" ▼", ""))

        if current_state == 0: # 정렬 안됨 -> 내림차순
            new_state = 1
            reverse = True
            arrow = " ▼"
        elif current_state == 1: # 내림차순 -> 오름차순
            new_state = 2
            reverse = False
            arrow = " ▲"
        else: # 오름차순 -> 정렬 취소
            new_state = 0
            reverse = False # 정렬 취소 시에는 순서 무의미
            arrow = ""

        self.sort_state = {c: 0 for c in self.result_tree["columns"]} # 모든 열 정렬 상태 초기화
        self.sort_state[col] = new_state

        # 현재 열 헤더에 화살표 추가
        current_text = self.result_tree.heading(col, "text")
        self.result_tree.heading(col, text=current_text.split(" ")[0] + arrow)

        # 데이터 가져오기
        data = []
        for item_id in self.result_tree.get_children():
            data.append((self.result_tree.item(item_id, 'values'), item_id))

        # 정렬
        if new_state != 0:
            col_index = self.result_tree["columns"].index(col)
            if col == "강수":
                data.sort(key=lambda x: self._get_round_value(x[0][col_index]), reverse=reverse)
            else:
                data.sort(key=lambda x: x[0][col_index], reverse=reverse)
            #data.sort(key=lambda x: x[0][col_index], reverse=reverse)

        # Treeview 업데이트
        for item_id in self.result_tree.get_children():
            self.result_tree.delete(item_id)

        for idx, (values, item_id) in enumerate(data):
            # '번호' 열을 현재 순서에 맞게 업데이트
            updated_values = list(values)
            updated_values[0] = idx + 1
            self.result_tree.insert('', 'end', values=updated_values)

    def _get_round_value(self, round_str):
        # Custom sort for "자유품새" rounds
        if '결선' in round_str:
            # Highest priority for ascending sort
            return (0, 0)
        elif '본선' in round_str:
            try:
                group_num = int(round_str.split('-')[1].replace('조', ''))
                # Second priority
                return (1, group_num)
            except (IndexError, ValueError):
                return (1, 0) # Fallback for "본선" without a group number
        elif '예선' in round_str:
            try:
                group_num = int(round_str.split('-')[1].replace('조', ''))
                # Third priority
                return (2, group_num)
            except (IndexError, ValueError):
                return (2, 0) # Fallback for "예선" without a group number

        # Sort logic for standard tournament rounds (e.g., "8", "4")
        try:
            # For ascending sort, smaller numbers come first.
            # For descending, larger numbers come first.
            # This is the natural integer order.
            return (3, int(round_str))
        except ValueError:
            # Fallback for any other string that doesn't fit the patterns above
            return (4, round_str)

    def export_results_to_excel(self):
        current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"경기번호계산_{current_time}.xlsx"

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                   initialfile=default_filename,
                                                   filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "경기 결과"

        # 헤더 추가
        headers = [self.result_tree.heading(col, "text").replace(" ▲", "").replace(" ▼", "") for col in self.result_tree["columns"]]
        sheet.append(headers)

        # 데이터 추가
        for item_id in self.result_tree.get_children():
            values = self.result_tree.item(item_id, 'values')
            sheet.append(values)

        try:
            workbook.save(file_path)
            tk.messagebox.showinfo("성공", f"결과가 {file_path}에 저장되었습니다.")
        except Exception as e:
            tk.messagebox.showerror("오류", f"파일 저장 중 오류가 발생했습니다: {e}")

    def add_row_with_data(self, data):
        row_frame = tk.Frame(self.main_frame)
        row_frame.pack(fill=tk.X, expand=True, pady=5)

        row_number = len(self.rows) + 1
        number_label = tk.Label(row_frame, text=str(row_number), width=5)
        number_label.pack(side=tk.LEFT, padx=5)

        labels = ["종목", "부", "체급", "참가인원"]
        entries = {}

        for i, label_text in enumerate(labels):
            entry = tk.Entry(row_frame, width=15)
            entry.pack(side=tk.LEFT, padx=5)
            if i < len(data):
                entry.insert(0, data[i])
            entries[label_text] = entry

        # '-' 버튼 추가
        remove_button_font = font.Font(family="Helvetica", size=12, weight="bold")
        remove_button = tk.Button(row_frame, text="-", bg="red", fg="white", font=remove_button_font,
                                  command=lambda r_frame=row_frame: self.remove_row(r_frame))
        remove_button.pack(side=tk.LEFT, padx=5)

        self.rows.append({"frame": row_frame, "entries": entries, "number_label": number_label, "remove_button": remove_button})

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    app = GameNumberCalculator(master=root)
    app.mainloop()
