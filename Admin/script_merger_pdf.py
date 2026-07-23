# -*- coding: utf-8 -*-
"""
merge_pdfs.py — PDF 파일 여러 개를 선택해서 순서를 정하고, 하나의 PDF로 합칩니다.

■ 사용법
    - python merge_pdfs.py 로 실행 (.pyw로 저장해서 실행해도 됨)
    - 파일 선택창이 뜨면 합칠 PDF들을 Ctrl(또는 Shift)로 여러 개 선택
    - 순서 조정 창이 뜸: "위로"/"아래로"로 병합 순서를 정하고, 필요 없는 파일은 "제거"로 뺄 수 있음
      (파일 선택창이 돌려주는 순서가 실제 클릭 순서와 다를 수 있어서, 순서 확인 단계를 넣었음)
    - "병합 시작"을 누르면 저장할 파일명을 물어보고, 진행 창이 뜬 뒤 결과 요약 창이 뜸

■ 요구사항
    - pip install pypdf
    (한/글이나 워드 같은 별도 프로그램 설치가 필요 없음 — Windows 전용도 아님)

■ 주의
    - 암호가 걸린 PDF는 읽을 수 없어 건너뜀 (실패 목록에 표시됨)
    - 페이지가 아주 많은 PDF들을 합치면 시간이 다소 걸릴 수 있음
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    tk.Tk().withdraw()
    messagebox.showerror(
        "오류",
        "pypdf가 설치되어 있지 않아요.\n\n"
        "명령 프롬프트에서 아래 명령을 실행한 뒤 다시 시도하세요:\n"
        "pip install pypdf"
    )
    raise SystemExit


def pick_order(root, files):
    """선택된 파일들의 병합 순서를 사용자가 조정할 수 있는 창을 띄우고,
    최종 순서(파일 경로 리스트)를 반환한다. 취소하면 None 반환."""
    win = tk.Toplevel(root)
    win.title("병합 순서 확인")
    win.geometry("480x360")

    tk.Label(win, text="병합될 순서입니다. 필요하면 위/아래로 옮기거나 제거하세요.").pack(pady=(10, 4))

    listbox = tk.Listbox(win, width=70, height=14)
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

    def remove():
        sel = listbox.curselection()
        if not sel:
            return
        i = sel[0]
        listbox.delete(i)
        order.pop(i)

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


def merge_pdfs(files, output_path, status_var, bar, progress_win):
    """files 순서대로 페이지를 이어붙여 output_path에 저장. (성공목록, 실패목록) 반환"""
    bar['maximum'] = len(files)
    done, failed = [], []

    writer = PdfWriter()
    for i, path in enumerate(files, 1):
        name = os.path.basename(path)
        status_var.set(f"({i}/{len(files)}) {name}")
        progress_win.update()
        try:
            reader = PdfReader(path)
            for page in reader.pages:
                writer.add_page(page)
            done.append(name)
        except Exception as e:
            failed.append((name, str(e)))

        bar['value'] = i
        progress_win.update()

    if done:
        with open(output_path, "wb") as f:
            writer.write(f)

    return done, failed


def main():
    root = tk.Tk()
    root.withdraw()

    files = filedialog.askopenfilenames(
        title="합칠 PDF 파일들 선택 (Ctrl 또는 Shift로 여러 개 선택)",
        filetypes=[("PDF 파일", "*.pdf")],
    )
    if not files:
        return

    order = pick_order(root, files)
    if not order:
        return

    output_path = filedialog.asksaveasfilename(
        title="병합된 PDF를 저장할 위치",
        defaultextension=".pdf",
        filetypes=[("PDF 파일", "*.pdf")],
        initialfile="merged.pdf",
    )
    if not output_path:
        return

    progress_win = tk.Toplevel(root)
    progress_win.title("PDF 병합 중...")
    progress_win.geometry("360x110")
    progress_win.resizable(False, False)

    status_var = tk.StringVar(value="시작 중...")
    tk.Label(progress_win, textvariable=status_var).pack(pady=(18, 6))
    bar = ttk.Progressbar(progress_win, length=320, mode='determinate')
    bar.pack(pady=6)
    progress_win.update()

    try:
        done, failed = merge_pdfs(order, output_path, status_var, bar, progress_win)
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