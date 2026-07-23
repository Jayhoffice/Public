# -*- coding: utf-8 -*-
"""
script_compressor_pdf.py — PDF 파일 여러 개를 선택하면(Ctrl 또는 Shift로 다중 선택),
Ghostscript를 이용해 압축된 PDF로 저장합니다.

■ 사용법
    - python script_compressor_pdf.py 로 실행 (.pyw로 저장해서 실행해도 됨)
    - 파일 선택창이 뜨면 압축할 PDF들을 Ctrl(또는 Shift)로 여러 개 선택
    - 압축 품질 선택창이 뜸: 화면용(최대 압축) / 이북용(균형, 기본값) / 인쇄용(고화질)
    - 진행 상황 창이 뜨고, 끝나면 결과(용량 변화 포함) 요약 창이 뜸
    - 압축된 파일은 원본 폴더에 "원본이름_compressed.pdf"로 저장됨 (원본은 건드리지 않음)

■ 요구사항
    - Ghostscript 설치 필요: https://ghostscript.com/releases/gsdnld.html
      (설치 시 "Add to PATH" 옵션을 꺼도 됨 — 이 스크립트가 흔한 설치 경로에서 자동으로 찾음)
    - 별도 pip 설치는 필요 없음 (표준 라이브러리만 사용)

■ 주의
    - 이미 텍스트 위주(스캔 이미지가 없는)인 PDF는 압축해도 용량이 크게 안 줄어들 수 있음.
      사진/스캔 이미지가 많이 들어간 PDF에서 효과가 큼.
    - 암호가 걸려 있거나 손상된 PDF는 실패 목록에 표시됨.
    - 동일한 이름의 압축 결과 파일이 이미 있으면 덮어씀.
    - Ghostscript에 넘기는 인자는 문자열로 조립하지 않고 리스트로 그대로 전달하므로,
      폴더/파일명에 공백이 있어도 별도 처리 없이 정상 동작함 — 예전에 Word COM
      자동화(doc_to_pdf.py)에서 겪었던 "경로를 URL로 오인식"하는 문제는 그쪽 COM
      구현 특유의 버그라 여기서는 애초에 발생 구조가 아님.
"""
import os
import glob
import shutil
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Ghostscript 압축 품질 프리셋: {키: (화면 표시용 라벨, Ghostscript PDFSETTINGS 값)}
QUALITY_PRESETS = {
    "screen": ("화면용 (최대 압축, 저해상도 72dpi)", "/screen"),
    "ebook": ("이북용 (균형, 150dpi) — 기본값", "/ebook"),
    "printer": ("인쇄용 (고화질, 300dpi)", "/printer"),
}
DEFAULT_QUALITY = "ebook"

HIDE_CONSOLE_WINDOW = True  # Ghostscript 실행 중 검은 콘솔 창이 안 보이게 함. 문제가 생기면 False로.


def find_ghostscript():
    """설치된 Ghostscript 실행 파일 경로를 찾는다. 못 찾으면 None."""
    # PATH에 등록되어 있는 경우 (Windows 64/32비트, macOS/Linux 공통 이름 순으로 확인)
    for name in ("gswin64c", "gswin32c", "gs"):
        path = shutil.which(name)
        if path:
            return path

    # PATH에 없는 경우, Windows 기본 설치 위치를 뒤진다 (버전 폴더명이 gs10.03.1 식으로 달라서 glob 사용)
    patterns = [
        r"C:\Program Files\gs\gs*\bin\gswin64c.exe",
        r"C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe",
    ]
    for pattern in patterns:
        matches = sorted(glob.glob(pattern), reverse=True)  # 버전 문자열 내림차순 = 최신 우선
        if matches:
            return matches[0]

    return None


def pick_quality(root):
    """압축 품질 프리셋을 고르는 창을 띄우고, 선택된 키를 반환한다. 취소하면 None."""
    win = tk.Toplevel(root)
    win.title("압축 품질 선택")
    win.geometry("380x220")
    win.resizable(False, False)

    tk.Label(win, text="압축 품질을 선택하세요.", pady=10).pack()

    choice = tk.StringVar(value=DEFAULT_QUALITY)
    for key, (label, _) in QUALITY_PRESETS.items():
        tk.Radiobutton(win, text=label, variable=choice, value=key).pack(anchor='w', padx=30, pady=2)

    result = {"key": None}

    def confirm():
        result["key"] = choice.get()
        win.destroy()

    btn_frame = tk.Frame(win)
    btn_frame.pack(pady=15)
    tk.Button(btn_frame, text="압축 시작", width=12, command=confirm).pack(side='left', padx=6)
    tk.Button(btn_frame, text="취소", width=12, command=win.destroy).pack(side='left', padx=6)

    win.grab_set()
    win.wait_window()
    return result["key"]


def compress_files(gs_exe, files, quality_key, status_var, bar, progress_win):
    """files를 압축해서 각각 '원본이름_compressed.pdf'로 저장.
    (성공목록[이름,이전용량,이후용량], 실패목록[이름,에러], 원본총용량, 압축후총용량) 반환"""
    bar['maximum'] = len(files)
    done, failed = [], []
    total_before, total_after = 0, 0

    pdfsettings = QUALITY_PRESETS[quality_key][1]

    # Windows에서 Ghostscript 실행 중 검은 콘솔 창이 깜빡이는 것을 막기 위한 설정
    startupinfo = None
    creationflags = 0
    if HIDE_CONSOLE_WINDOW and os.name == 'nt':
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        creationflags = subprocess.CREATE_NO_WINDOW

    for i, src in enumerate(files, 1):
        name = os.path.basename(src)
        status_var.set(f"({i}/{len(files)}) {name}")
        progress_win.update()

        base, ext = os.path.splitext(src)
        out_path = base + "_compressed" + ext

        try:
            # 인자를 리스트로 그대로 넘기므로(쉘 문자열 조립 없음) 경로/파일명에
            # 공백이 있어도 subprocess가 알아서 올바르게 따옴표 처리를 해준다.
            result = subprocess.run(
                [
                    gs_exe,
                    "-sDEVICE=pdfwrite",
                    "-dCompatibilityLevel=1.4",
                    f"-dPDFSETTINGS={pdfsettings}",
                    "-dNOPAUSE", "-dQUIET", "-dBATCH",
                    f"-sOutputFile={out_path}",
                    src,
                ],
                capture_output=True, text=True,
                startupinfo=startupinfo, creationflags=creationflags,
            )
            if result.returncode != 0 or not os.path.exists(out_path):
                raise RuntimeError(result.stderr.strip()[-200:] or "알 수 없는 오류")

            before = os.path.getsize(src)
            after = os.path.getsize(out_path)
            total_before += before
            total_after += after
            done.append((name, before, after))
        except Exception as e:
            failed.append((name, str(e)))

        bar['value'] = i
        progress_win.update()

    return done, failed, total_before, total_after


def main():
    root = tk.Tk()
    root.withdraw()

    gs_exe = find_ghostscript()
    if not gs_exe:
        messagebox.showerror(
            "Ghostscript를 찾을 수 없음",
            "Ghostscript가 설치되어 있지 않은 것 같아요.\n\n"
            "https://ghostscript.com/releases/gsdnld.html 에서 설치한 뒤 다시 실행해주세요.\n"
            "(설치할 때 PATH 등록을 안 해도, 기본 설치 경로면 이 스크립트가 자동으로 찾습니다.)"
        )
        return

    files = filedialog.askopenfilenames(
        title="압축할 PDF 파일들 선택 (Ctrl 또는 Shift로 여러 개 선택)",
        filetypes=[("PDF 파일", "*.pdf")],
    )
    if not files:
        return

    quality_key = pick_quality(root)
    if not quality_key:
        return

    progress_win = tk.Toplevel(root)
    progress_win.title("PDF 압축 중...")
    progress_win.geometry("360x110")
    progress_win.resizable(False, False)

    status_var = tk.StringVar(value="시작 중...")
    tk.Label(progress_win, textvariable=status_var).pack(pady=(18, 6))
    bar = ttk.Progressbar(progress_win, length=320, mode='determinate')
    bar.pack(pady=6)
    progress_win.update()

    try:
        done, failed, total_before, total_after = compress_files(
            gs_exe, files, quality_key, status_var, bar, progress_win
        )
    except Exception as e:
        progress_win.destroy()
        messagebox.showerror("오류", f"압축 중 문제가 발생했어요:\n{e}")
        return

    progress_win.destroy()

    if not done:
        msg = "압축에 성공한 파일이 없어요."
    else:
        def fmt(n):
            return f"{n / 1024 / 1024:.1f}MB" if n >= 1024 * 1024 else f"{n / 1024:.0f}KB"

        pct = (1 - total_after / total_before) * 100 if total_before else 0
        msg = (
            f"총 {len(files)}개 중 {len(done)}개 압축 완료.\n"
            f"용량: {fmt(total_before)} → {fmt(total_after)} ({pct:.0f}% 감소)\n"
            f"압축된 파일은 원본과 같은 폴더에 '_compressed'가 붙어서 저장됐어요."
        )
    if failed:
        msg += "\n\n실패한 파일:\n" + "\n".join(f"- {n}: {e}" for n, e in failed)
    messagebox.showinfo("압축 결과", msg)


if __name__ == '__main__':
    main()
