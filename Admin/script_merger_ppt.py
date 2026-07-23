# -*- coding: utf-8 -*-
"""
merge_powerpoints_keep_background.py
— PPT/PPTX 파일 여러 개를 선택해서 순서를 정하고, 하나의 PowerPoint 파일로 합칩니다.

■ 핵심 개선점
    - 기존 Slides.InsertFromFile 방식은 슬라이드는 잘 합쳐지지만,
      원본 슬라이드 마스터/레이아웃에 들어 있던 배경 이미지나 마스터 도형이
      대상 프레젠테이션의 디자인으로 바뀌면서 사라질 수 있음.
    - 이를 보완하기 위해 두 번째 파일부터는 각 원본 슬라이드의
      '배경 + 마스터/레이아웃 영역'만 PNG로 임시 캡처한 뒤,
      병합된 슬라이드의 가장 뒤쪽에 배경 그림으로 삽입함.
    - 슬라이드 본문의 텍스트/도형/차트 등은 기존 PowerPoint 객체 상태로 유지됨.

■ 사용법
    - python merge_powerpoints_keep_background.py 로 실행 (.pyw로 저장해서 실행해도 됨)
    - 파일 선택창이 뜨면 합칠 PPT/PPTX들을 Ctrl(또는 Shift)로 여러 개 선택
    - 순서 조정 창에서 위/아래 이동 또는 제거
    - '병합 시작' → 저장 위치 선택 → 진행률 표시 → 결과 요약

■ 요구사항
    - Windows
    - Microsoft PowerPoint 데스크톱 앱 설치 필요
    - pip install pywin32

■ 주의
    - .ppt(구형 형식)까지 지원하기 위해 PowerPoint COM 자동화를 사용함
    - 첫 번째 정상 파일이 결과 프레젠테이션의 기반이 됨
    - 두 번째 파일부터는 배경 보존을 위해 슬라이드별 PNG가 결과 파일에 포함되므로
      기존 버전보다 결과 PPTX 용량이 커질 수 있음
    - 배경 이미지는 시각적으로 보존되지만, 슬라이드 마스터 자체가 완전히 합쳐지는 것은 아님
    - 서로 다른 화면 비율(예: 4:3, 16:9)을 섞으면 일부 개체 배치가 달라질 수 있음
    - 읽을 수 없거나 손상된 파일은 건너뛰고 실패 목록에 표시함
"""

import os
import tempfile
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

# Office MsoTriState / MsoZOrderCmd 상수
MSO_FALSE = 0
MSO_TRUE = -1
MSO_SEND_TO_BACK = 1

# 배경 캡처 해상도: 16:9 기준 1920x1080 수준
BACKGROUND_EXPORT_WIDTH = 1920


def pick_order(root, files):
    """선택된 파일들의 병합 순서를 사용자가 조정할 수 있는 창을 띄우고,
    최종 순서(파일 경로 리스트)를 반환한다. 취소하면 None 반환."""
    win = tk.Toplevel(root)
    win.title("병합 순서 확인")
    win.geometry("520x380")

    tk.Label(
        win,
        text="병합될 순서입니다. 필요하면 위/아래로 옮기거나 제거하세요."
    ).pack(pady=(10, 4))

    listbox = tk.Listbox(win, width=76, height=15)
    for f in files:
        listbox.insert(tk.END, os.path.basename(f))
    listbox.pack(padx=10, pady=4, fill="both", expand=True)

    order = list(files)

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

        if listbox.size() > 0:
            new_i = min(i, listbox.size() - 1)
            listbox.selection_set(new_i)
            listbox.activate(new_i)

    btn_frame = tk.Frame(win)
    btn_frame.pack(pady=4)
    tk.Button(btn_frame, text="위로", width=8, command=lambda: move(-1)).pack(side="left", padx=4)
    tk.Button(btn_frame, text="아래로", width=8, command=lambda: move(1)).pack(side="left", padx=4)
    tk.Button(btn_frame, text="제거", width=8, command=remove).pack(side="left", padx=4)

    result = {"order": None}

    def confirm():
        if not order:
            messagebox.showwarning("알림", "합칠 파일이 없어요.")
            return
        result["order"] = list(order)
        win.destroy()

    action_frame = tk.Frame(win)
    action_frame.pack(pady=10)
    tk.Button(action_frame, text="병합 시작", width=12, command=confirm).pack(side="left", padx=6)
    tk.Button(action_frame, text="취소", width=12, command=win.destroy).pack(side="left", padx=6)

    win.grab_set()
    win.wait_window()
    return result["order"]


def get_save_format(output_path):
    """저장 경로 확장자에 맞는 PowerPoint SaveAs 형식 상수를 반환."""
    ext = os.path.splitext(output_path)[1].lower()
    if ext == ".ppt":
        return PP_SAVE_AS_PPT
    return PP_SAVE_AS_PPTX


def same_path(path1, path2):
    """Windows 경로의 대소문자 차이 등을 무시하고 같은 파일인지 비교."""
    return os.path.normcase(os.path.abspath(path1)) == os.path.normcase(os.path.abspath(path2))


def safe_filename(text):
    """임시 파일명으로 사용할 수 있도록 특수문자를 단순 치환."""
    invalid = '<>:"/\\|?*'
    result = text
    for ch in invalid:
        result = result.replace(ch, "_")
    return result


def export_background_only(source_pres, slide_index, temp_dir, prefix):
    """
    원본 슬라이드를 임시 복제한 뒤 슬라이드 자체의 모든 Shapes를 삭제하고,
    남아 있는 '슬라이드 배경 + 마스터/레이아웃 배경 요소'만 PNG로 Export한다.

    반환값: 생성된 PNG 경로

    ※ 원본 슬라이드는 변경하지 않는다.
    """
    original_slide = source_pres.Slides(slide_index)
    temp_slide = None

    try:
        duplicated_range = original_slide.Duplicate()
        temp_slide = duplicated_range.Item(1)

        # 슬라이드 레벨의 객체(텍스트, 이미지, 차트 등)를 모두 제거한다.
        # 마스터/레이아웃에 있는 객체와 실제 Background는 Shapes 컬렉션에 포함되지 않으므로
        # 이 상태로 Export하면 사실상 배경 레이어만 남는다.
        while temp_slide.Shapes.Count > 0:
            temp_slide.Shapes.Item(1).Delete()

        slide_width = float(source_pres.PageSetup.SlideWidth)
        slide_height = float(source_pres.PageSetup.SlideHeight)

        export_width = BACKGROUND_EXPORT_WIDTH
        export_height = max(1, int(round(export_width * slide_height / slide_width)))

        filename = f"{safe_filename(prefix)}_slide_{slide_index:04d}_bg.png"
        png_path = os.path.join(temp_dir, filename)

        temp_slide.Export(png_path, "PNG", export_width, export_height)

        if not os.path.exists(png_path):
            raise RuntimeError("배경 PNG 생성에 실패했습니다.")

        return png_path

    finally:
        if temp_slide is not None:
            try:
                temp_slide.Delete()
            except Exception:
                pass


def add_background_image(dest_slide, png_path, slide_width, slide_height):
    """
    PNG를 슬라이드 전체 크기로 삽입하고 가장 뒤로 보낸다.
    삽입된 슬라이드의 본문 객체는 그 위에 그대로 남아 편집 가능하다.
    """
    bg = dest_slide.Shapes.AddPicture(
        png_path,
        MSO_FALSE,   # LinkToFile=False
        MSO_TRUE,    # SaveWithDocument=True
        0,
        0,
        slide_width,
        slide_height,
    )

    try:
        bg.Name = "__MERGED_SOURCE_BACKGROUND__"
    except Exception:
        pass

    bg.ZOrder(MSO_SEND_TO_BACK)


def merge_powerpoints(files, output_path, status_var, bar, progress_win):
    """
    files 순서대로 슬라이드를 이어붙여 output_path에 저장.
    (성공목록, 실패목록) 반환.

    첫 번째 정상 파일은 결과 파일의 기반으로 그대로 사용한다.
    이후 파일은:
      1) 원본 슬라이드별 배경 레이어를 PNG로 캡처
      2) Slides.InsertFromFile로 슬라이드 삽입
      3) 캡처한 배경 PNG를 각 삽입 슬라이드의 맨 뒤에 추가
    """
    bar["maximum"] = len(files)
    done, failed = [], []

    powerpoint = None
    merged = None
    com_initialized = False

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")

        with tempfile.TemporaryDirectory(prefix="ppt_merge_bg_") as temp_dir:
            for i, path in enumerate(files, 1):
                name = os.path.basename(path)
                status_var.set(f"({i}/{len(files)}) {name}")
                progress_win.update()

                try:
                    abs_path = os.path.abspath(path)

                    if merged is None:
                        # 첫 번째 정상 파일은 그대로 결과 파일의 기반으로 사용한다.
                        candidate = powerpoint.Presentations.Open(
                            abs_path,
                            False,  # ReadOnly=False
                            False,  # Untitled=False
                            False,  # WithWindow=False
                        )

                        try:
                            save_format = get_save_format(output_path)
                            candidate.SaveAs(os.path.abspath(output_path), save_format)
                            merged = candidate
                            done.append(name)
                        except Exception:
                            candidate.Close()
                            raise

                    else:
                        source = None
                        background_files = []

                        try:
                            # 배경 캡처용으로 원본 프레젠테이션을 연다.
                            # 저장하지 않고 닫으므로 원본 파일 자체는 변경되지 않는다.
                            source = powerpoint.Presentations.Open(
                                abs_path,
                                True,   # ReadOnly=True
                                False,  # Untitled=False
                                False,  # WithWindow=False
                            )

                            source_slide_count = source.Slides.Count
                            if source_slide_count <= 0:
                                raise RuntimeError("슬라이드가 없습니다.")

                            # 1. 원본 배경을 먼저 모두 캡처
                            for slide_no in range(1, source_slide_count + 1):
                                status_var.set(
                                    f"({i}/{len(files)}) {name} - 배경 보존 {slide_no}/{source_slide_count}"
                                )
                                progress_win.update()

                                bg_path = export_background_only(
                                    source,
                                    slide_no,
                                    temp_dir,
                                    f"file_{i}",
                                )
                                background_files.append(bg_path)

                            # 2. 기존 방식으로 슬라이드 본문 삽입
                            insert_after = merged.Slides.Count
                            inserted_count = merged.Slides.InsertFromFile(
                                abs_path,
                                insert_after,
                            )

                            if inserted_count <= 0:
                                raise RuntimeError("삽입된 슬라이드가 없습니다.")

                            if inserted_count != len(background_files):
                                raise RuntimeError(
                                    f"삽입 슬라이드 수({inserted_count})와 배경 수"
                                    f"({len(background_files)})가 일치하지 않습니다."
                                )

                            # 3. 삽입된 각 슬라이드 뒤에 원본 배경 PNG 추가
                            merged_width = float(merged.PageSetup.SlideWidth)
                            merged_height = float(merged.PageSetup.SlideHeight)

                            for offset, bg_path in enumerate(background_files, start=1):
                                dest_slide_index = insert_after + offset
                                dest_slide = merged.Slides(dest_slide_index)

                                status_var.set(
                                    f"({i}/{len(files)}) {name} - 배경 적용 {offset}/{inserted_count}"
                                )
                                progress_win.update()

                                add_background_image(
                                    dest_slide,
                                    bg_path,
                                    merged_width,
                                    merged_height,
                                )

                            done.append(name)

                        finally:
                            if source is not None:
                                try:
                                    source.Close()
                                except Exception:
                                    pass

                except Exception as e:
                    failed.append((name, str(e)))

                bar["value"] = i
                progress_win.update()

            if merged is not None and done:
                merged.Save()

    finally:
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

    if any(same_path(output_path, path) for path in order):
        messagebox.showerror(
            "오류",
            "저장할 파일 경로가 선택한 원본 파일 중 하나와 같습니다.\n\n"
            "원본이 아닌 다른 이름이나 위치로 저장하세요."
        )
        return

    progress_win = tk.Toplevel(root)
    progress_win.title("PowerPoint 병합 중...")
    progress_win.geometry("430x120")
    progress_win.resizable(False, False)

    status_var = tk.StringVar(value="PowerPoint 시작 중...")
    tk.Label(progress_win, textvariable=status_var).pack(pady=(18, 6))
    bar = ttk.Progressbar(progress_win, length=380, mode="determinate")
    bar.pack(pady=6)
    progress_win.update()

    try:
        done, failed = merge_powerpoints(
            order,
            output_path,
            status_var,
            bar,
            progress_win,
        )
    except Exception as e:
        progress_win.destroy()
        messagebox.showerror("오류", f"병합 중 문제가 발생했어요:\n{e}")
        return

    progress_win.destroy()

    if not done:
        messagebox.showerror(
            "병합 결과",
            "합칠 수 있는 파일이 없어서 저장하지 못했어요."
        )
        return

    msg = (
        f"총 {len(order)}개 중 {len(done)}개 병합 완료.\n"
        f"저장 위치: {output_path}"
    )

    if failed:
        msg += "\n\n실패한 파일(건너뜀):\n" + "\n".join(
            f"- {n}: {e}" for n, e in failed
        )

    messagebox.showinfo("병합 결과", msg)


if __name__ == "__main__":
    main()