# -*- coding: utf-8 -*-
r"""
folder_name_watcher.py — 내가 파일명에 어떤 "분류명"을 썼는지 잊지 않게, 화면 우측 상단에
항상 띄워두는 작은 상태창.

찾는 패턴: 1차분류_2차분류_...  (첫 번째와 두 번째 밑줄 사이 부분들, 단 1차분류는
    반드시 문자로 시작해야 함)
    예) 수업_물리학세미나1_26-1학기_수업및평가계획.hwp -> 1차="수업", 2차="물리학세미나1"
        평가_물리학세미나1_26_1학기_기말시험.hwp        -> 1차="평가", 2차="물리학세미나1"
        20260716_backup.txt                          -> (숫자로 시작해서 제외됨)
        정리_회의록.hwp                                -> (밑줄이 하나뿐이라 2차분류가 없어서 제외됨)

감시할 폴더는 실행할 때 직접 고릅니다(여러 개 가능). 각 폴더는 하위 폴더까지
전부(재귀적으로) 훑어서, (1차분류, 2차분류) 조합별로 파일이 몇 개인지 개수를 세어
1차분류 합계가 많은 순서(내림차순)로 그룹을 정렬하고, 같은 1차분류 안에서는
2차분류 개수가 많은 순서로 표에 보여줍니다. 같은 1차분류가 여러 줄에 걸치면
그 칸을 세로로 병합해서 가운데 정렬로 한 번만 보여줍니다.
목록이 창보다 길면 마우스 휠이나 스크롤바로 볼 수 있습니다.

■ 사용법
    - python folder_name_watcher.py 로 처음 실행하면 폴더 선택 창이 뜹니다. '폴더 추가'로
      감시할 폴더를 원하는 만큼 추가한 뒤 '확인'을 누르면 상태창이 뜨고, 고른 폴더
      목록이 저장됩니다.
    - 다음에 실행할 때는 저장된 폴더 목록을 자동으로 불러와서, 선택 창 없이 바로
      상태창이 뜹니다. 감시 폴더를 바꾸고 싶으면 상단 바의 '📂' 버튼을 누르세요.
    - 자동 새로고침은 없음 — 상단 바 '↻'를 눌러야 그때그때 최신 상태로 갱신됨.
      'x'는 종료.
    - 상단 바를 마우스로 눌러서 드래그하면 위치 이동.
    - 창 우측 하단 모서리(⤡)를 드래그하면 크기 조절.
    - 표의 행을 더블클릭하면, 그 행의 "1차_2차" 문자열을 검색어로 넣어서 Everything이 열립니다.
      (예: 수업 | 물리학세미나1 행을 더블클릭 -> Everything에서 "수업_물리학세미나1" 검색)
      -> 아래 EVERYTHING_EXE 경로를 자신의 설치 위치에 맞게 고쳐야 동작합니다.
    - 저장된 폴더 목록은 %USERPROFILE%\.category_watcher_folders.json 에 저장됨.
      이 파일을 지우면 다음 실행 때 다시 처음처럼 선택 창부터 뜸.
    - 자동 시작/시작 메뉴 등록 기능은 없음 — 필요할 때마다 이 파일을 직접 실행.

■ 자원 사용량 안내
    - 자동 새로고침이 없어서 평소에는 완전히 대기 상태라 CPU를 거의 안 씀.
      '↻'를 누른 순간에만 선택한 폴더들을 하위 폴더까지 전부 훑습니다.
      파일이 몇백 개 수준이면 그 순간에도 가볍지만(1000분의 몇 초 수준),
      파일/폴더 수가 아주 많은 폴더(예: 용량이 큰 클라우드 동기화 폴더)를
      지정하면 새로고침할 때마다 그만큼 시간이 더 걸릴 수 있습니다.
    - 메모리도 tkinter 창 하나 띄워두는 수준(대략 20~40MB)이라, 오래된 컴퓨터에서도
      브라우저 탭 하나보다 훨씬 가볍습니다.
"""
import os
import re
import json
import subprocess
import threading
import queue
from collections import Counter
from itertools import groupby
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont

# 1차분류_2차분류_... 패턴
# 밑줄(_)이 두 개 이상 있는 파일명이면, 첫 번째 밑줄 앞부분을 1차분류로,
# 첫 번째와 두 번째 밑줄 사이를 2차분류로 봄.
# 단, 1차분류는 반드시 "문자(한글 또는 영문)로 시작"해야 함 — 숫자로 시작하는 이름
# (예: 20260716_backup.txt 같은 시스템/기타 파일)이 잘못 걸리는 것을 막기 위함.
# 밑줄이 하나뿐이라 2차분류가 없는 파일(예: 정리_회의록.hwp)은 대상에서 제외됨.
CATEGORY_PATTERN = re.compile(r'^([A-Za-z가-힣][^_]*)_([^_]*)_')

# 폴더 선택 창을 열 때 기본 시작 위치로만 씀 (실제 감시 폴더는 사용자가 고름)
DEFAULT_PICK_DIR = os.path.join(os.path.expanduser('~'), 'Desktop')

# 감시 폴더 목록을 저장해두는 파일 (다음 실행 때 자동으로 불러오기 위함)
CONFIG_PATH = os.path.join(os.path.expanduser('~'), '.category_watcher_folders.json')

# Everything(voidtools) 실행 파일 경로. 자신의 설치 위치에 맞게 고칠 것.
# (바탕화면/시작메뉴의 Everything 바로가기를 마우스 우클릭 -> 속성 -> '대상'에서 확인 가능)
EVERYTHING_EXE = r'C:\Program Files\Everything\Everything.exe'

MIN_WIDTH, MIN_HEIGHT = 190, 120
DEFAULT_WIDTH, DEFAULT_HEIGHT = 260, 280


def load_saved_folders():
    """이전에 저장해 둔 폴더 목록을 불러옴. 저장된 게 없거나 읽을 수 없으면 None."""
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return None
    return data if isinstance(data, list) and data else None


def save_folders(roots):
    """선택한 폴더 목록을 다음 실행 때 불러올 수 있도록 저장.
    저장에 실패해도 프로그램 동작에는 지장 없음(다음 실행 때 선택 창이 다시 뜰 뿐)."""
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(roots, f, ensure_ascii=False)
    except OSError:
        pass


def select_folders(initial=None, parent=None):
    """감시할 폴더를 고르는 선택 창을 띄움. '폴더 추가'로 여러 개(2~3개 등)를
    담을 수 있고, '확인'을 누르면 선택된 폴더 경로 리스트를 반환.
    폴더를 하나도 고르지 않고 창을 닫으면 None을 반환.
    initial: 미리 채워둘 폴더 목록(폴더 변경 시 현재 목록을 넘겨줄 때 사용).
    parent: 상태창에서 '📂'로 다시 열 때, 그 위젯의 root(Toplevel로 띄워서
    Tk() 인스턴스를 중복 생성하지 않기 위함)."""
    picker = tk.Toplevel(parent) if parent is not None else tk.Tk()
    picker.title('감시할 폴더 선택')
    picker.geometry('420x320')

    tk.Label(picker, text='분류명을 감시할 폴더를 선택하세요 (여러 개 가능)').pack(pady=(10, 4))

    listbox = tk.Listbox(picker, selectmode='extended')
    listbox.pack(fill='both', expand=True, padx=10, pady=4)
    for folder in (initial or []):
        listbox.insert('end', folder)

    def add_folder():
        folder = filedialog.askdirectory(title='폴더 선택', initialdir=DEFAULT_PICK_DIR)
        if folder and folder not in listbox.get(0, 'end'):
            listbox.insert('end', folder)

    def remove_selected():
        for i in reversed(listbox.curselection()):
            listbox.delete(i)

    btn_frame = tk.Frame(picker)
    btn_frame.pack(fill='x', padx=10, pady=4)
    tk.Button(btn_frame, text='폴더 추가', command=add_folder).pack(side='left')
    tk.Button(btn_frame, text='선택 삭제', command=remove_selected).pack(side='left', padx=6)

    result = []

    def confirm():
        result.extend(listbox.get(0, 'end'))
        picker.destroy()

    tk.Button(picker, text='확인', command=confirm).pack(pady=(4, 10))

    if parent is not None:
        picker.transient(parent)
        picker.grab_set()
        parent.wait_window(picker)
    else:
        picker.mainloop()
    return result if result else None


def get_category_counts(roots):
    """여러 폴더의 파일/폴더 이름을 하위 폴더까지 전부(재귀적으로) 읽어서,
    (1차분류, 2차분류)별 개수를 폴더 전체 합산으로 (개수 내림차순으로) 반환.
    접근할 수 없는 폴더는 건너뛰고, 그 사유는 errors 리스트에 모아서 함께 반환함."""
    counter = Counter()
    errors = []

    def on_walk_error(exc):
        errors.append(f"폴더에 접근 권한이 없어요:\n{exc.filename}")

    for root_dir in roots:
        if not os.path.isdir(root_dir):
            errors.append(f"폴더를 찾을 수 없어요:\n{root_dir}")
            continue

        for _dirpath, dirnames, filenames in os.walk(root_dir, onerror=on_walk_error):
            for name in dirnames + filenames:
                m = CATEGORY_PATTERN.match(name)
                if m:
                    counter[(m.group(1), m.group(2))] += 1

    # 1차분류별 합계를 구해서, 합계가 큰 1차분류 그룹이 위로 오도록 정렬하고
    # (같은 1차분류는 항상 붙어 있어야 표에서 병합해서 보여줄 수 있으므로,
    # 합계가 같아 순서가 애매할 때는 1차분류명으로 묶어줌)
    # 같은 1차분류 안에서는 2차분류 개수가 많은 순서로 정렬함.
    cat1_totals = Counter()
    for (cat1, _cat2), count in counter.items():
        cat1_totals[cat1] += count

    items = sorted(
        counter.items(),
        key=lambda kv: (-cat1_totals[kv[0][0]], kv[0][0], -kv[1], kv[0][1])
    )
    return items, errors


def get_category_files(roots):
    """여러 폴더의 파일/폴더 이름을 하위 폴더까지 전부(재귀적으로) 읽어서,
    (1차분류, 2차분류)별로 매칭된 이름들을 모아서 반환.
    반환값은 get_category_counts와 같은 순서(개수 내림차순)로 정렬된
    [((1차, 2차), [이름, ...]), ...] 리스트와 errors 리스트."""
    groups = {}
    errors = []

    def on_walk_error(exc):
        errors.append(f"폴더에 접근 권한이 없어요:\n{exc.filename}")

    for root_dir in roots:
        if not os.path.isdir(root_dir):
            errors.append(f"폴더를 찾을 수 없어요:\n{root_dir}")
            continue

        for _dirpath, dirnames, filenames in os.walk(root_dir, onerror=on_walk_error):
            for name in dirnames + filenames:
                m = CATEGORY_PATTERN.match(name)
                if m:
                    groups.setdefault((m.group(1), m.group(2)), []).append(name)

    items = sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0]))
    return items, errors


class CategoryWidget:
    def __init__(self, roots):
        self.roots = roots
        self.root = tk.Tk()
        self.root.overrideredirect(True)   # 제목표시줄 없는 작은 창
        self.root.attributes('-topmost', True)
        self.root.attributes('-alpha', 0.92)
        self.root.configure(bg='#222222')
        self.root.minsize(MIN_WIDTH, MIN_HEIGHT)

        x = self.root.winfo_screenwidth() - DEFAULT_WIDTH - 20
        self.root.geometry(f'{DEFAULT_WIDTH}x{DEFAULT_HEIGHT}+{x}+20')

        # 상단 바 (드래그 + 폴더변경 + 새로고침 + 닫기 버튼)
        bar = tk.Frame(self.root, bg='#333333', height=24)
        bar.pack(fill='x', side='top')
        tk.Label(bar, text='📁 분류명', bg='#333333', fg='white',
                 font=('맑은 고딕', 9, 'bold')).pack(side='left', padx=6)
        tk.Button(bar, text='✕', bg='#333333', fg='white', bd=0,
                  command=self.root.destroy, font=('맑은 고딕', 9)).pack(side='right', padx=4)
        tk.Button(bar, text='↻', bg='#333333', fg='white', bd=0,
                  command=self.refresh, font=('맑은 고딕', 9)).pack(side='right')
        tk.Button(bar, text='📂', bg='#333333', fg='white', bd=0,
                  command=self.change_folders, font=('맑은 고딕', 9)).pack(side='right')
        tk.Button(bar, text='📄', bg='#333333', fg='white', bd=0,
                  command=self.export_file_list, font=('맑은 고딕', 9)).pack(side='right')
        bar.bind('<Button-1>', self.start_move)
        bar.bind('<B1-Motion>', self.do_move)

        # 상태 메시지(빈 목록/에러) 표시용
        self.status_label = tk.Label(self.root, text='', bg='#222222', fg='#888888',
                                      font=('맑은 고딕', 9), justify='left', anchor='nw', wraplength=220)

        # 본문: 1차분류 | 2차분류 | 개수 표 (스크롤 가능)
        # ttk.Treeview는 셀 병합을 지원하지 않아서, 같은 1차분류를 세로로 병합해
        # 가운데 정렬로 보여주기 위해 Frame + Label 그리드로 직접 표를 그림.
        # 열 구성: 0=1차분류, 1=폭조절 손잡이, 2=2차분류, 3=폭조절 손잡이, 4=개수
        body = tk.Frame(self.root, bg='#222222')
        body.pack(fill='both', expand=True, padx=6, pady=4)
        self.body = body

        # 사용자가 손잡이로 직접 조절한 폭(픽셀). None이면 아직 조절 안 한 상태 -> 자동 폭 사용.
        # [1차분류 폭, 2차분류 폭] (개수 칸은 항상 자동 폭)
        self.col_widths = [None, None]
        # 실제 글자 폭 기준 최소 폭(내용이 잘리지 않는 최소치). _render에서 매번 갱신됨.
        self.min_col_widths = [0, 0, 0]
        self.HANDLE_WIDTH = 4

        # 표 머리글 (고정, 스크롤되지 않음)
        header = tk.Frame(body, bg='#333333')
        header.pack(fill='x', side='top')
        tk.Label(header, text='1차', bg='#333333', fg='white',
                 font=('맑은 고딕', 9, 'bold'), anchor='center').grid(row=0, column=0, sticky='ew', padx=(4, 2))

        handle1 = tk.Frame(header, bg='#555555', width=self.HANDLE_WIDTH, cursor='sb_h_double_arrow')
        handle1.grid(row=0, column=1, sticky='ns')
        handle1.bind('<Button-1>', lambda e: self._start_col_resize(e, 0))
        handle1.bind('<B1-Motion>', lambda e: self._do_col_resize(e, 0))

        tk.Label(header, text='2차', bg='#333333', fg='white',
                 font=('맑은 고딕', 9, 'bold'), anchor='w').grid(row=0, column=2, sticky='w')

        handle2 = tk.Frame(header, bg='#555555', width=self.HANDLE_WIDTH, cursor='sb_h_double_arrow')
        handle2.grid(row=0, column=3, sticky='ns')
        handle2.bind('<Button-1>', lambda e: self._start_col_resize(e, 1))
        handle2.bind('<B1-Motion>', lambda e: self._do_col_resize(e, 1))

        tk.Label(header, text='개수', bg='#333333', fg='white',
                 font=('맑은 고딕', 9, 'bold'), anchor='w').grid(row=0, column=4, sticky='w', padx=(4, 4))
        # 열 너비는 _render/_apply_column_widths에서 실제 글자 폭에 맞춰 header/inner 양쪽에
        # 동일한 minsize로 정해줌. weight를 전부 0으로 둬서, 남는 공간이 어느 한 칸으로
        # 자동으로 밀려들어가지 않게 함(그래야 손잡이로 좁힐 때 진짜로 좁아짐).
        header.grid_columnconfigure(0, weight=0)
        header.grid_columnconfigure(1, weight=0)
        header.grid_columnconfigure(2, weight=0)
        header.grid_columnconfigure(3, weight=0)
        header.grid_columnconfigure(4, weight=0)
        self.header = header

        # 표 몸통 (스크롤 가능한 캔버스 위에 실제 행들을 그림)
        table_area = tk.Frame(body, bg='#222222')
        table_area.pack(fill='both', expand=True, side='top')

        self.canvas = tk.Canvas(table_area, bg='#222222', highlightthickness=0)
        scrollbar = ttk.Scrollbar(table_area, orient='vertical', command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.inner = tk.Frame(self.canvas, bg='#222222')
        self.inner_window = self.canvas.create_window((0, 0), window=self.inner, anchor='nw')
        # 손잡이가 있는 1,3번 열은 헤더와 폭만 맞추면 되므로 내용은 없음(빈 칸)
        self.inner.grid_columnconfigure(0, weight=0)
        self.inner.grid_columnconfigure(1, weight=0, minsize=self.HANDLE_WIDTH)
        self.inner.grid_columnconfigure(2, weight=0)
        self.inner.grid_columnconfigure(3, weight=0, minsize=self.HANDLE_WIDTH)
        self.inner.grid_columnconfigure(4, weight=0)



        # 내용 크기가 바뀌면 스크롤 범위 갱신, 창 크기가 바뀌면 표 너비를 맞춤
        self.inner.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.bind('<Configure>', lambda e: self.canvas.itemconfig(self.inner_window, width=e.width))
        # 마우스 휠 스크롤 (Windows에서는 보통 기본 지원되지만, 명시적으로 한 번 더 걸어둠)
        self.canvas.bind('<MouseWheel>', self._on_mousewheel)

        # 우측 하단 크기 조절 손잡이
        grip = tk.Label(self.root, text='⤡', bg='#333333', fg='#999999',
                         font=('맑은 고딕', 9), cursor='sizing')
        grip.place(relx=1.0, rely=1.0, anchor='se', width=16, height=16)
        grip.bind('<Button-1>', self.start_resize)
        grip.bind('<B1-Motion>', self.do_resize)

        self._queue = queue.Queue()
        self._poll_queue()  # 큐를 주기적으로 확인하는 루프 시작
        self.refresh()

    # --- 창 이동 ---
    def start_move(self, event):
        self._dx, self._dy = event.x, event.y

    def do_move(self, event):
        x = self.root.winfo_pointerx() - self._dx
        y = self.root.winfo_pointery() - self._dy
        self.root.geometry(f'+{x}+{y}')

    # --- 창 크기 조절 ---
    def start_resize(self, event):
        self._start_w = self.root.winfo_width()
        self._start_h = self.root.winfo_height()
        self._start_x = event.x_root
        self._start_y = event.y_root

    def do_resize(self, event):
        dw = event.x_root - self._start_x
        dh = event.y_root - self._start_y
        new_w = max(MIN_WIDTH, self._start_w + dw)
        new_h = max(MIN_HEIGHT, self._start_h + dh)
        self.root.geometry(f'{new_w}x{new_h}')

    # --- 표 칸 폭 조절 (헤더의 손잡이를 드래그) ---
    def _start_col_resize(self, event, col_index):
        self._resize_start_x = event.x_root
        current = self.col_widths[col_index]
        self._resize_start_width = current if current is not None else self.min_col_widths[col_index]

    def _do_col_resize(self, event, col_index):
        delta = event.x_root - self._resize_start_x
        new_width = max(self.min_col_widths[col_index], self._resize_start_width + delta)
        self.col_widths[col_index] = new_width
        self._apply_column_widths()

    def _apply_column_widths(self):
        """1차분류(0번 칸)/2차분류(2번 칸)/개수(4번 칸)의 실제 폭을 header와 inner
        양쪽에 똑같이 적용해서 정렬을 맞춤. 사용자가 손잡이로 직접 조절한 폭이 있으면
        그 값을, 없으면 자동으로 계산된 최소 폭을 사용함(둘 다 최소 폭보다 작아지진 않음)."""
        cat1_w = max(self.col_widths[0] or 0, self.min_col_widths[0])
        cat2_w = max(self.col_widths[1] or 0, self.min_col_widths[1])
        count_w = self.min_col_widths[2]
        for frame in (self.header, self.inner):
            frame.grid_columnconfigure(0, minsize=cat1_w)
            frame.grid_columnconfigure(2, minsize=cat2_w)
            frame.grid_columnconfigure(4, minsize=count_w)

    # --- 마우스 휠 스크롤 ---
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

    def open_in_everything(self, query):
        try:
            subprocess.Popen([EVERYTHING_EXE, '-s', query])
        except FileNotFoundError:
            messagebox.showerror(
                'Everything을 찾을 수 없어요',
                f'다음 경로에 Everything.exe가 없어요:\n{EVERYTHING_EXE}\n\n'
                '파일 상단의 EVERYTHING_EXE 값을 실제 설치 경로로 고쳐주세요.'
            )

    # --- 데이터 새로고침 (별도 스레드) ---
    def refresh(self):
        threading.Thread(target=self._fetch_in_background, daemon=True).start()

    # --- 폴더 변경 ---
    def change_folders(self):
        new_roots = select_folders(initial=self.roots, parent=self.root)
        if new_roots:
            self.roots = new_roots
            save_folders(self.roots)
            self.refresh()

    # --- 카테고리로 잡힌 파일명 txt 일괄 출력 ---
    def export_file_list(self):
        items, errors = get_category_files(self.roots)

        if not items:
            text = '\n\n'.join(errors) if errors else '내보낼 분류명이 없어요.'
            messagebox.showinfo('내보낼 파일이 없어요', text)
            return

        path = filedialog.asksaveasfilename(
            title='파일명 목록 저장',
            defaultextension='.txt',
            filetypes=[('텍스트 파일', '*.txt')],
            initialfile='분류_파일목록.txt',
        )
        if not path:
            return

        lines = []
        for (cat1, cat2), names in items:
            lines.append(f'[{cat1}_{cat2}] ({len(names)}개)')
            lines.extend(sorted(names))
            lines.append('')

        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines))
        except OSError as exc:
            messagebox.showerror('저장 실패', f'파일을 저장하지 못했어요:\n{exc}')
            return

        messagebox.showinfo('완료', f'{len(items)}개 분류의 파일명을 저장했어요:\n{path}')

    def _fetch_in_background(self):
        items, errors = get_category_counts(self.roots)
        self._queue.put((items, errors))

    def _poll_queue(self):
        """tkinter는 다른 스레드에서 직접 위젯을 건드리면 위험하므로,
        메인 스레드에서 주기적으로 큐를 확인하는 방식으로 안전하게 결과를 반영"""
        try:
            items, errors = self._queue.get_nowait()
        except queue.Empty:
            pass
        else:
            self._render(items, errors)
        self.root.after(200, self._poll_queue)

    # --- 표 렌더링 ---
    def _render(self, items, errors):
        if not items:
            text = '\n\n'.join(errors) if errors else '(아직 분류명 없음)'
            self.status_label.config(text=text, fg='#ff8888' if errors else '#888888')
            self.status_label.pack(fill='both', expand=True, padx=8, pady=6)
            self.body.pack_forget()
            return

        self.status_label.pack_forget()
        if not self.body.winfo_ismapped():
            self.body.pack(fill='both', expand=True, padx=6, pady=4)

        for child in self.inner.winfo_children():
            child.destroy()

        # items는 이미 1차분류별로 붙어서 정렬되어 있으므로(get_category_counts 참고),
        # 같은 1차분류끼리 묶어서 그 칸을 세로로 병합(rowspan)하고 가운데 정렬함.
        grouped = []
        for cat1, group in groupby(items, key=lambda kv: kv[0][0]):
            subitems = [(cat2, count) for (_cat1, cat2), count in group]
            grouped.append((cat1, subitems))

        # 열 너비는 글자 수 고정이 아니라 실제 글자 폭을 재서 정함 -> 1차분류(수업/평가/행정처럼
        # 짧은 이름)는 저절로 좁아지고 2차분류(더 긴 이름)는 저절로 넓어지며, 창 크기와 무관하게
        # 글자가 잘리는 일이 없음. header와 inner 양쪽에 같은 minsize를 줘서 열을 맞춤.
        # (사용자가 헤더의 손잡이로 직접 넓혀놓은 폭이 있으면 이 최소 폭 위에서 유지됨)
        body_font = tkfont.Font(family='맑은 고딕', size=10)
        header_font = tkfont.Font(family='맑은 고딕', size=9, weight='bold')
        cell_pad = 14

        self.min_col_widths[0] = max([header_font.measure('1차')] +
                                      [body_font.measure(cat1) for cat1, _ in grouped]) + cell_pad
        self.min_col_widths[1] = max([header_font.measure('2차')] +
                                      [body_font.measure(cat2) for _cat1, sub in grouped for cat2, _ in sub]) + cell_pad
        self.min_col_widths[2] = max([header_font.measure('개수')] +
                                      [body_font.measure(str(c)) for _cat1, sub in grouped for _cat2, c in sub]) + cell_pad
        self._apply_column_widths()

        # 1차분류 그룹이 어디서 바뀌는지 한눈에 보이도록, 그룹마다 배경색을 뚜렷하게 번갈아 표시
        band_colors = ('#1a1a1a', '#3a3a3a')

        row = 0
        for group_index, (cat1, subitems) in enumerate(grouped):
            span = len(subitems)
            bg = band_colors[group_index % 2]

            cat1_label = tk.Label(self.inner, text=cat1, bg=bg, fg='#dddddd',
                                   font=('맑은 고딕', 10), anchor='center')
            cat1_label.grid(row=row, column=0, rowspan=span, sticky='nsew', padx=(4, 2), pady=1)
            cat1_label.bind('<MouseWheel>', self._on_mousewheel)
            # 병합된 1차분류 칸을 더블클릭하면 1차분류명만으로 Everything 검색
            cat1_label.bind('<Double-1>', lambda e, c1=cat1: self.open_in_everything(c1))

            for cat2, count in subitems:
                cat2_label = tk.Label(self.inner, text=cat2, bg=bg, fg='#dddddd',
                                       font=('맑은 고딕', 10), anchor='w')
                cat2_label.grid(row=row, column=2, sticky='nsew', pady=1)

                count_label = tk.Label(self.inner, text=str(count), bg=bg, fg='#dddddd',
                                        font=('맑은 고딕', 10), anchor='w')
                count_label.grid(row=row, column=4, sticky='nsew', padx=(4, 4), pady=1)

                # 행 더블클릭 -> 해당 (1차_2차)로 Everything 검색
                for widget in (cat2_label, count_label):
                    widget.bind('<MouseWheel>', self._on_mousewheel)
                    widget.bind('<Double-1>', lambda e, c1=cat1, c2=cat2: self.open_in_everything(f'{c1}_{c2}'))

                row += 1

        # 처음 실행하거나 새 폴더를 봐서 내용이 넓어졌을 때, 개수 칸이 창 밖으로 밀려서
        # 안 보이는 일이 없도록 필요한 만큼만 창 폭을 자동으로 넓힘(좁히지는 않음 ->
        # 사용자가 이미 늘려놓은 창 크기는 그대로 유지됨).
        self.root.update_idletasks()
        scrollbar_and_margin = 40
        required_width = self.header.winfo_reqwidth() + scrollbar_and_margin
        if required_width > self.root.winfo_width():
            self.root.geometry(
                f'{required_width}x{self.root.winfo_height()}'
                f'+{self.root.winfo_x()}+{self.root.winfo_y()}'
            )

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    saved_roots = load_saved_folders()
    if saved_roots:
        CategoryWidget(saved_roots).run()
    else:
        chosen_roots = select_folders()
        if chosen_roots:
            save_folders(chosen_roots)
            CategoryWidget(chosen_roots).run()