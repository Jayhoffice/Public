# -*- coding: utf-8 -*-
"""
hwp_to_pdf.py — hwp/hwpx 파일 여러 개를 선택하면(Ctrl 또는 Shift로 다중 선택),
한/글(한컴오피스) 자체 PDF 변환 기능으로 pdf로 저장합니다.
(프린터로 인쇄하는 방식이 아니라, 한/글의 "다른 이름으로 저장 > PDF" 기능을 그대로 이용)

■ 사용법
    - python hwp_to_pdf.py 로 실행 (.pyw로 저장해서 실행해도 됨)
    - 파일 선택창이 뜨면 변환할 hwp/hwpx 파일들을 Ctrl(또는 Shift)로 여러 개 선택
    - 진행 상황 창이 뜨고, 끝나면 결과(성공/실패 개수) 요약 창이 뜸
    - pdf는 원본과 같은 폴더에 같은 파일명으로 저장됨 (동일한 이름의 pdf가 있으면 덮어씀)

■ 요구사항
    - Windows + 한/글(한컴오피스) 설치
    - pip install pyhwpx

■ 주의
    - 변환 중에는 한/글 창이 화면에 나타났다 사라지는 것이 반복될 수 있음(정상 동작).
      HIDE_HWP_WINDOW = True 로 두면 창을 숨겨서 이 깜빡임을 없앨 수 있음.
    - 파일이 아주 많으면(수백 개) 그만큼 시간이 걸림 — 진행 창에 몇 번째 파일인지 표시됨.
    - [미검증 예방 조치] Word 자동화에서 폴더 경로에 공백이 있으면 파일을 못 찾는
      것으로 오작동하는 문제가 확인된 적이 있어서, 한/글도 같은 문제가 있을 가능성에
      대비해 Open에 넘기는 경로를 짧은(8.3) 경로로 바꿔서 사용하도록 해뒀음. 다만 이건
      한/글에서 실제로 확인된 문제는 아니라서, 혹시 이 조치 이후 오히려 안 되던 게
      생기면 아래 주석 표시된 부분을 되돌리면 됨.
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from pyhwpx import Hwp
    import win32api
except ImportError:
    tk.Tk().withdraw()
    messagebox.showerror(
        "오류",
        "pyhwpx가 설치되어 있지 않아요.\n\n"
        "명령 프롬프트에서 아래 명령을 실행한 뒤 다시 시도하세요:\n"
        "pip install pyhwpx"
    )
    raise SystemExit

HIDE_HWP_WINDOW = True  # 변환 중 한/글 창을 안 보이게 함(깜빡임 방지). 문제가 생기면 False로.


def convert_files(files, status_var, bar, progress_win):
    """선택된 hwp/hwpx 파일들을 pdf로 변환. (성공목록, 실패목록) 반환"""
    bar['maximum'] = len(files)
    done, failed = [], []

    hwp = Hwp()  # 실행 시 한/글 보안모듈이 자동으로 등록되어, 저장 중 확인창이 뜨지 않음
    if HIDE_HWP_WINDOW:
        hwp.XHwpWindows.Item(0).Visible = False

    try:
        for i, src in enumerate(files, 1):
            name = os.path.basename(src)
            status_var.set(f"({i}/{len(files)}) {name}")
            progress_win.update()

            pdf_path = os.path.splitext(src)[0] + '.pdf'

            # [미검증 예방 조치] Open에 넘길 경로만 짧은(8.3, 공백 없는) 경로로 바꿔서
            # 쓴다. 변환 실패 시(예: 해당 드라이브에서 8.3 이름 생성이 꺼져있는 경우)
            # 원래 경로를 그대로 쓴다. 저장할 pdf 경로(pdf_path)는 원래 이름을 유지해야
            # 하므로 건드리지 않는다.
            open_path = src
            try:
                open_path = win32api.GetShortPathName(src)
            except Exception:
                pass

            opened = False
            try:
                if not hwp.Open(open_path):
                    raise RuntimeError("파일 열기 실패")
                opened = True

                # 주의: SaveAs(path, "PDF")는 예전에 "현재 쪽만" 등으로 인쇄했던
                # 설정이 남아있으면 그 설정이 그대로 PDF 변환에도 적용되는 결함이
                # 있음(한/글 자체 버그). 그래서 SaveAs 대신 인쇄 액션(PrintToPDF)의
                # Range 값을 "0 = 문서전체"로 직접 못박아서 변환한다.
                # pset.Range = 0 처럼 속성으로 직접 접근하면, 설치된 한/글 버전의
                # 타입 라이브러리에 따라 일부 필드(FileName 등)가 없다는 에러가 날 수
                # 있어서, 버전에 상관없이 항상 되는 SetItem()으로 지정한다.
                pset = hwp.HParameterSet.HPrint
                hwp.HAction.GetDefault("PrintToPDF", pset.HSet)
                pset.HSet.SetItem("Range", 0)  # 0 = 문서 전체 (연결된 문서 포함)
                pset.HSet.SetItem("FileName", pdf_path)
                if not hwp.HAction.Execute("PrintToPDF", pset.HSet):
                    raise RuntimeError("PDF 저장 실패")

                done.append(name)
            except Exception as e:
                failed.append((name, str(e)))
            finally:
                # 성공/실패와 상관없이, 문서를 열었다면 반드시 닫는다.
                # (여기서 안 닫으면 실패한 파일의 문서가 계속 열린 채로 남아있게 되고,
                #  그 상태로 다음 파일들을 계속 열다 보면 한/글이 불안정해져서
                #  뒤쪽 파일들에서 연쇄적으로 에러가 나기 쉬움 — 원래 코드의 버그였음)
                if opened:
                    try:
                        hwp.Clear(1)  # 1 = 변경사항 버리고 문서 닫기
                    except Exception:
                        pass

            bar['value'] = i
            progress_win.update()
    finally:
        hwp.Quit()  # 성공/실패와 무관하게 한/글 프로세스는 반드시 종료

    return done, failed


def main():
    root = tk.Tk()
    root.withdraw()

    files = filedialog.askopenfilenames(
        title="변환할 hwp/hwpx 파일 선택 (Ctrl 또는 Shift로 여러 개 선택)",
        filetypes=[("한/글 문서", "*.hwp;*.hwpx")],
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