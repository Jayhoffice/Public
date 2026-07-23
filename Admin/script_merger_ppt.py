# -*- coding: utf-8 -*-
"""
merge_powerpoints.py — PPT/PPTX 파일 여러 개를 선택해서 순서를 정하고, 하나의 PowerPoint 파일로 합칩니다.

■ 사용법
    - python merge_powerpoints.py 로 실행 (.pyw로 저장해서 실행해도 됨)
    - 파일 선택창이 뜨면 합칠 PPT/PPTX들을 Ctrl(또는 Shift)로 여러 개 선택
    - 순서 조정 창이 뜸: "위로"/"아래로"로 병합 순서를 정하고, 필요 없는 파일은 "제거"로 뺄 수 있음
      (파일 선택창이 돌려주는 순서가 실제 클릭 순서와 다를 수 있어서, 순서 확인 단계를 넣었음)
    - "병합 시작"을 누르면 저장할 파일명을 물어보고, 진행 창이 뜬 뒤 결과 요약 창이 뜸

■ 요구사항
    - Windows
    - Microsoft PowerPoint 데스크톱 앱 설치 필요
    - pip install pywin32

■ 주의
    - .ppt(구형 형식)까지 지원하기 위해 PowerPoint 자체의 COM 자동화 기능을 사용함
    - 첫 번째로 정상 처리된 파일이 병합 결과의 기본 프레젠테이션이 됨
      (슬라이드 크기/페이지 비율도 기본적으로 이 파일을 기준으로 함)
    - 서로 다른 화면 비율(예: 4:3, 16:9)의 파일을 섞으면 일부 개체 배치가 달라질 수 있음
    - 읽을 수 없거나 손상된 파일은 건너뛰고 실패 목록에 표시함
    - 저장 형식은 .pptx 또는 .ppt 중 선택 가능하지만, 특별한 이유가 없다면 .pptx 사용을 권장함
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pythoncom
    import win32com.client
except ImportError:
    tk.Tk().withdraw()
    messagebox.showerror(
        "오류",
        "pywin32가 설치되어 있지 않아요.\n\n"
        "명령 프롬프트에서 아래 명령을 실행한 뒤 다시 시도하세요:\n"
        "pip install pywin32"
    )
    raise SystemExit


# PowerPoint PpSaveAsFileType 상수
PP_SAVE_AS_PPT = 1       # ppSaveAsPresentation (.ppt)
PP_SAVE_AS_PPTX = 24     # ppSaveAsOpenXMLPresentation (.pptx)


def pick_order(root, files):
    """선택된 파일들의 병합 순서를 사용자가 조정할 수 있는 창을 띄우고,
    최종 순서(파일 경로 리스트)를 반환한다. 취소하면 None 반환."""
    win = tk.Toplevel(root)
    win.title("병합 순서 확인")
    win.geometry("520x380")

    tk.Label(win, text="병합될 순서입니다. 필요하면 위/아래로 옮기거나 제거하세요.").pack(pady=(10, 4))

    listbox = tk.Listbox(win, width=76, height=15)
    for f in files:
        listbox.insert(tk.END, os.path.basename(f))
    listbox.pack(padx=10, pady=4, fill='both', expand=True)

    order = list(files)  # listbox 항목과 같은 순서로 유지되는 실제 경로 리스트

    def move(offset):
        sel = listbox.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + offset
        if j < 0 or j >= listbox.size():
            return
        order[i], order[j] = order[j], order[i]
        text_i, text_j = listbox.get(i), listbox.get(j)
        listbox.delete(i)
        listbox.insert(i, text_j)
        listbox.delete(j)
        listbox.insert(j, text_i)
        listbox.selection_set(j)
        listbox.activate(j)

    def remove():
        sel = listbox.curselection()
        if not sel:
            return
        i = sel[0]
        listbox.delete(i)
        order.pop(i)

        # 제거 후 가능한 한 가까운 항목을 다시 선택
        if listbox.size() > 0:
            new_i = min(i, listbox.size() - 1)
            listbox.selection_set(new_i)
            listbox.activate(new_i)

    btn_frame = tk.Frame(win)
    btn_frame.pack(pady=4)
    tk.Button(btn_frame, text="위로", width=8, command=lambda: move(-1)).pack(side='left', padx=4)
    tk.Button(btn_frame, text="아래로", width=8, command=lambda: move(1)).pack(side='left', padx=4)
    tk.Button(btn_frame, text="제거", width=8, command=remove).pack(side='left', padx=4)

    result = {"order": None}

    def confirm():
        if not order:
            messagebox.showwarning("알림", "합칠 파일이 없어요.")
            return
        result["order"] = list(order)
        win.destroy()

    action_frame = tk.Frame(win)
    action_frame.pack(pady=10)
    tk.Button(action_frame, text="병합 시작", width=12, command=confirm).pack(side='left', padx=6)
    tk.Button(action_frame, text="취소", width=12, command=win.destroy).pack(side='left', padx=6)

    win.grab_set()
    win.wait_window()
    return result["order"]


def get_save_format(output_path):
    """저장 경로의 확장자에 맞는 PowerPoint SaveAs 형식 상수를 반환한다."""
    ext = os.path.splitext(output_path)[1].lower()
    if ext == ".ppt":
        return PP_SAVE_AS_PPT
    return PP_SAVE_AS_PPTX


def same_path(path1, path2):
    """Windows 경로의 대소문자 차이 등을 무시하고 같은 파일인지 비교한다."""
    return os.path.normcase(os.path.abspath(path1)) == os.path.normcase(os.path.abspath(path2))


def merge_powerpoints(files, output_path, status_var, bar, progress_win):
    """files 순서대로 슬라이드를 이어붙여 output_path에 저장.
    (성공목록, 실패목록) 반환.

    첫 번째로 정상적으로 열린 프레젠테이션을 결과 파일의 기반으로 사용하고,
    이후 파일들은 PowerPoint의 Slides.InsertFromFile로 뒤에 이어 붙인다.
    """
    bar['maximum'] = len(files)
    done, failed = [], []

    powerpoint = None
    merged = None
    com_initialized = False

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        # 별도의 PowerPoint 프로세스를 열어 기존에 사용 중인 PowerPoint와 충돌을 줄임
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")

        for i, path in enumerate(files, 1):
            name = os.path.basename(path)
            status_var.set(f"({i}/{len(files)}) {name}")
            progress_win.update()

            try:
                abs_path = os.path.abspath(path)

                if merged is None:
                    # 첫 번째 정상 파일을 기반 프레젠테이션으로 사용
                    # ReadOnly=False, Untitled=False, WithWindow=False
                    candidate = powerpoint.Presentations.Open(abs_path, False, False, False)

                    try:
                        save_format = get_save_format(output_path)
                        candidate.SaveAs(os.path.abspath(output_path), save_format)
                        merged = candidate
                        done.append(name)
                    except Exception:
                        candidate.Close()
                        raise
                else:
                    # 현재 마지막 슬라이드 뒤에 해당 파일의 모든 슬라이드를 삽입
                    insert_after = merged.Slides.Count
                    inserted_count = merged.Slides.InsertFromFile(abs_path, insert_after)

                    if inserted_count <= 0:
                        raise RuntimeError("삽입된 슬라이드가 없습니다.")

                    done.append(name)

            except Exception as e:
                failed.append((name, str(e)))

            bar['value'] = i
            progress_win.update()

        if merged is not None and done:
            merged.Save()

    finally:
        # 오류가 나더라도 PowerPoint 프로세스가 백그라운드에 남지 않도록 정리
        if merged is not None:
            try:
                merged.Close()
            except Exception:
                pass

        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass

        if com_initialized:
            pythoncom.CoUninitialize()

    return done, failed


def main():
    root = tk.Tk()
    root.withdraw()

    files = filedialog.askopenfilenames(
        title="합칠 PPT/PPTX 파일들 선택 (Ctrl 또는 Shift로 여러 개 선택)",
        filetypes=[
            ("PowerPoint 파일", "*.pptx *.ppt"),
            ("PPTX 파일", "*.pptx"),
            ("PPT 파일", "*.ppt"),
        ],
    )
    if not files:
        return

    order = pick_order(root, files)
    if not order:
        return

    output_path = filedialog.asksaveasfilename(
        title="병합된 PowerPoint를 저장할 위치",
        defaultextension=".pptx",
        filetypes=[
            ("PowerPoint 프레젠테이션 (*.pptx)", "*.pptx"),
            ("PowerPoint 97-2003 프레젠테이션 (*.ppt)", "*.ppt"),
        ],
        initialfile="merged.pptx",
    )
    if not output_path:
        return

    # 입력 파일 자체를 결과 파일로 덮어쓰는 실수를 방지
    if any(same_path(output_path, path) for path in order):
        messagebox.showerror(
            "오류",
            "저장할 파일 경로가 선택한 원본 파일 중 하나와 같습니다.\n\n"
            "원본이 아닌 다른 이름이나 위치로 저장하세요."
        )
        return

    progress_win = tk.Toplevel(root)
    progress_win.title("PowerPoint 병합 중...")
    progress_win.geometry("400x110")
    progress_win.resizable(False, False)

    status_var = tk.StringVar(value="PowerPoint 시작 중...")
    tk.Label(progress_win, textvariable=status_var).pack(pady=(18, 6))
    bar = ttk.Progressbar(progress_win, length=350, mode='determinate')
    bar.pack(pady=6)
    progress_win.update()

    try:
        done, failed = merge_powerpoints(order, output_path, status_var, bar, progress_win)
    except Exception as e:
        progress_win.destroy()
        messagebox.showerror("오류", f"병합 중 문제가 발생했어요:\n{e}")
        return

    progress_win.destroy()

    if not done:
        messagebox.showerror("병합 결과", "합칠 수 있는 파일이 없어서 저장하지 못했어요.")
        return

    msg = f"총 {len(order)}개 중 {len(done)}개 병합 완료.\n저장 위치: {output_path}"
    if failed:
        msg += "\n\n실패한 파일(건너뜀):\n" + "\n".join(f"- {n}: {e}" for n, e in failed)
    messagebox.showinfo("병합 결과", msg)


if __name__ == '__main__':
    main()
