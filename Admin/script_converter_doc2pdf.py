# -*- coding: utf-8 -*-
"""
doc_to_pdf.py — doc/docx 파일 여러 개를 선택하면(Ctrl 또는 Shift로 다중 선택),
Microsoft Word 자체 PDF 변환 기능으로 pdf로 저장합니다.
(Word의 "다른 이름으로 저장 > PDF" 대신, 전용 PDF 내보내기 기능인
 ExportAsFixedFormat을 그대로 이용 — SaveAs보다 안정적)

■ 사용법
    - python doc_to_pdf.py 로 실행 (.pyw로 저장해서 실행해도 됨)
    - 파일 선택창이 뜨면 변환할 doc/docx 파일들을 Ctrl(또는 Shift)로 여러 개 선택
    - 진행 상황 창이 뜨고, 끝나면 결과(성공/실패 개수) 요약 창이 뜸
    - pdf는 원본과 같은 폴더에 같은 파일명으로 저장됨 (동일한 이름의 pdf가 있으면 덮어씀)

■ 요구사항
    - Windows + Microsoft Word 설치
    - pip install pywin32

■ 주의
    - 변환 중 Word 창이 화면에 나타났다 사라지는 것이 반복될 수 있음(정상 동작).
      HIDE_WORD_WINDOW = True 로 두면 창을 숨겨서 이 깜빡임을 없앨 수 있음.
    - 인터넷에서 받은 파일(다운로드 표시가 붙은 파일)은 Word가 "보호된 보기"로 열면서
      자동화가 멈출 수 있음. 이 경우는 이 스크립트로는 못 잡음 — Word 신뢰 센터 설정에서
      "보호된 보기"를 끄거나, 파일 속성에서 차단 해제를 해야 함.
    - 파일이 아주 많으면(수백 개) 그만큼 시간이 걸림 — 진행 창에 몇 번째 파일인지 표시됨.
    - 파일이 들어있는 폴더 경로에 공백이 있으면 Word의 COM 자동화가 파일을 못 찾는
      것으로 오작동하는 경우가 있어(경로를 URL로 잘못 해석하는 것으로 추정), Open에
      넘기는 경로만 짧은(8.3) 경로로 변환해 이 문제를 회피함. 저장되는 pdf 파일명은
      원래 이름 그대로 유지됨(짧은 경로는 여는 동작에만 사용, 저장 경로는 원래 경로 사용).
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import win32com.client as win32
    import win32api
except ImportError:
    tk.Tk().withdraw()
    messagebox.showerror(
        "오류",
        "pywin32가 설치되어 있지 않아요.\n\n"
        "명령 프롬프트에서 아래 명령을 실행한 뒤 다시 시도하세요:\n"
        "pip install pywin32"
    )
    raise SystemExit

HIDE_WORD_WINDOW = True  # 변환 중 Word 창을 안 보이게 함(깜빡임 방지). 문제가 생기면 False로.

WD_EXPORT_FORMAT_PDF = 17     # wdExportFormatPDF
WD_EXPORT_ALL_DOCUMENT = 0    # wdExportAllDocument (문서 전체)
WD_ALERTS_NONE = 0            # wdAlertsNone


def convert_files(files, status_var, bar, progress_win):
    """선택된 doc/docx 파일들을 pdf로 변환. (성공목록, 실패목록) 반환"""
    bar['maximum'] = len(files)
    done, failed = [], []

    word = win32.Dispatch('Word.Application')
    word.Visible = not HIDE_WORD_WINDOW
    word.DisplayAlerts = WD_ALERTS_NONE  # 서식 호환성 등 모든 확인창이 뜨지 않게 함

    try:
        for i, src in enumerate(files, 1):
            name = os.path.basename(src)
            status_var.set(f"({i}/{len(files)}) {name}")
            progress_win.update()

            pdf_path = os.path.splitext(src)[0] + '.pdf'

            # 폴더 경로에 공백이 있으면 Word COM 자동화가 파일을 못 찾는다고 오작동하는
            # 경우가 있어서, Open에 넘길 경로만 짧은(8.3, 공백 없는) 경로로 바꿔서 쓴다.
            # 변환 실패 시(예: 해당 드라이브에서 8.3 이름 생성이 꺼져있는 경우) 원래
            # 경로를 그대로 쓴다. 저장할 pdf 경로(pdf_path)는 원래 이름을 유지해야 하므로
            # 건드리지 않는다.
            open_path = src
            try:
                open_path = win32api.GetShortPathName(src)
            except Exception:
                pass

            doc = None
            try:
                doc = word.Documents.Open(
                    open_path,
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                )
                # SaveAs(FileFormat=17)보다 전용 PDF 내보내기 함수가 더 안정적이라
                # ExportAsFixedFormat을 사용. Range를 문서 전체로 명시해서, 혹시
                # 남아있을 수 있는 이전 인쇄 범위 설정의 영향을 받지 않게 한다.
                doc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=WD_EXPORT_FORMAT_PDF,
                    Range=WD_EXPORT_ALL_DOCUMENT,
                )
                done.append(name)
            except Exception as e:
                failed.append((name, str(e)))
            finally:
                # 성공/실패와 상관없이, 문서를 열었다면 반드시 닫는다.
                # (안 닫으면 실패한 파일의 문서가 계속 열린 채로 남아있게 되고,
                #  그 상태로 다음 파일들을 계속 열다 보면 Word가 불안정해져서
                #  뒤쪽 파일들에서 연쇄적으로 에러가 나기 쉬움)
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass

            bar['value'] = i
            progress_win.update()
    finally:
        word.Quit()  # 성공/실패와 무관하게 Word 프로세스는 반드시 종료

    return done, failed


def main():
    root = tk.Tk()
    root.withdraw()

    files = filedialog.askopenfilenames(
        title="변환할 doc/docx 파일 선택 (Ctrl 또는 Shift로 여러 개 선택)",
        filetypes=[("Word 문서", "*.doc;*.docx")],
    )
    if not files:
        return

    progress_win = tk.Toplevel(root)
    progress_win.title("PDF로 변환 중...")
    progress_win.geometry("360x110")
    progress_win.resizable(False, False)

    status_var = tk.StringVar(value="시작 중...")
    tk.Label(progress_win, textvariable=status_var).pack(pady=(18, 6))
    bar = ttk.Progressbar(progress_win, length=320, mode='determinate')
    bar.pack(pady=6)
    progress_win.update()

    try:
        done, failed = convert_files(files, status_var, bar, progress_win)
    except Exception as e:
        progress_win.destroy()
        messagebox.showerror("오류", f"변환 중 문제가 발생했어요:\n{e}")
        return

    progress_win.destroy()

    msg = f"총 {len(files)}개 중 {len(done)}개 변환 완료.\npdf는 원본과 같은 폴더에 저장됐어요."
    if failed:
        msg += "\n\n실패한 파일:\n" + "\n".join(f"- {n}: {e}" for n, e in failed)
    messagebox.showinfo("변환 결과", msg)


if __name__ == '__main__':
    main()