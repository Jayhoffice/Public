# -*- coding: utf-8 -*-
"""
hwp(x)-to-md.py — HWP / HWPX -> Obsidian용 Markdown 통합 변환기

사용법 1) 그냥 실행 (권장) - 파일 선택 창이 뜸, 여러 개 선택 가능(Ctrl/Shift+클릭)
    python hwp(x)-to-md.py

사용법 2) 파일 경로를 직접 지정 (자동화/스크립트용)
    python hwp(x)-to-md.py 입력파일.hwp
    python hwp(x)-to-md.py 입력파일.hwpx -o 원하는출력.md
    python hwp(x)-to-md.py 입력파일.hwp --no-hangul   (한글 프로그램 없이, .hwp는 텍스트만 추출)

동작 원칙 (구조 파악 목적, 완벽 재현 목적 아님):
- .hwp 파일이 선택되면: 설치된 한글(한컴오피스) 프로그램을 pyhwpx로 배경 조종해서
  같은 폴더에 .hwpx로 먼저 저장한 뒤, 그 hwpx를 아래 파이프라인으로 변환함.
  이때 만들어지는 .hwpx는 "md로 변환하기 위한 중간 산출물"일 뿐이라, md 변환이
  끝나면(성공/실패 상관없이) 바로 자동 삭제됨 — 폴더에 남지 않음.
  (한글 프로그램이 없거나 --no-hangul 옵션을 쓰면, .hwp는 텍스트만 추출)
- 수식이 없는 문서 -> 그냥 텍스트만 쭉 추출 (제목 기호 #, 표 등 아무 구조도 만들지 않음)
- 수식이 있는 문서 -> 다음까지 포함한 전체 파이프라인
    · "페이지 틀 용도로 쓰인 표"(칸이 1개뿐인 행)는 표 대신 줄글로 풀어씀
    · 수식은 LaTeX(\\dfrac, \\langle 등)으로 변환, 대괄호(\\[ \\]) delimiter는 쓰지 않음
    · 원본 텍스트 내용은 임의로 바꾸지 않음 (수식 표기법 변환 제외)

출력 파일은 원본과 같은 폴더에 같은 이름으로 저장됨 (예: 문서.hwp -> 문서.md)
같은 이름의 .md 파일이 이미 있으면 물어보지 않고 덮어씀 — 주의.

■ 처음 설정하는 법 (Windows, 최초 1회만 하면 됨)
    1) VS Code(또는 아무 터미널)에서 파이썬이 설치되어 있는지 확인:
           python --version
       버전 숫자가 안 나오고 에러가 뜨면, 아래 명령으로 설치:
           winget install Python.Python.3.12
       설치 후 터미널(VS Code)을 완전히 껐다가 다시 켜고 python --version 재확인.

    2) 필요한 라이브러리 설치 (터미널에 한 줄씩 입력, 순서대로):
           pip install pyhwp
           pip install pyhwpx
       (tkinter는 파이썬에 보통 이미 포함되어 있어서 따로 설치 안 해도 되는 경우가 대부분입니다.
        혹시 "No module named tkinter" 에러가 뜨면, 파이썬을 처음 설치할 때 옵션에서
        tcl/tk 관련 항목이 빠졌을 수 있으니 python.org에서 설치파일을 다시 받아 재설치하세요.)

    3) 이 파일(hwp(x)-to-md.py)을 하나의 폴더에 저장.

    4) 터미널에서 그 폴더로 이동 후 실행:
           cd 이 파일이 있는 폴더 경로
           python hwp(x)-to-md.py
       파일 선택 창이 뜨면 정상 설정된 것입니다.

    ※ pip install 중 "externally-managed-environment" 같은 에러가 뜨면
       (드물지만 일부 환경에서 발생):
           pip install pyhwp --break-system-packages
           pip install pyhwpx --break-system-packages
       처럼 뒤에 --break-system-packages 를 붙여서 다시 시도하세요.

■ 실행 환경 요구사항
    - .hwp 파일을 변환하려면 "Windows PC에 한글(한컴오피스) 프로그램이 실제로 설치되어
      있어야 합니다." pyhwpx가 그 프로그램을 배경에서 직접 조종하는 방식이라, 프로그램 자체가
      없으면 이 기능은 동작하지 않습니다 (이 경우 .hwpx만 다루거나 --no-hangul 옵션을 쓰세요).
    - .hwpx 파일만 다룰 거라면 한글 프로그램 설치는 필요 없습니다.

■ 사용자가 알아둬야 할 주의사항
    1) .hwp 변환 중에는 한글 프로그램 창이 화면에 실제로 잠깐씩 떴다 사라집니다.
       변환이 끝날 때까지는 그 창을 직접 닫거나 클릭하지 말고 기다려주세요
       (자동 작업 중에 손대면 오류가 날 수 있습니다).
    2) 파일을 여러 개 선택하면 한글 프로그램을 한 번만 켜서 순서대로 재사용합니다
       (파일마다 매번 새로 켜지 않아 더 빠름). 처리 도중 파일 하나가 실패해도
       나머지 파일은 계속 진행되고, 끝나고 나서 어떤 파일이 성공/실패했는지 목록으로 보여줍니다.
    3) 원본 .hwp/.hwpx 파일 자체는 절대 수정하지 않습니다. 오직 같은 폴더에
       .md 파일을 새로 만들 뿐입니다 (단, .hwp의 경우 위에서 설명한 임시 .hwpx가
       아주 잠깐 생겼다가 자동으로 지워집니다).
    4) 이 변환기는 "내용을 알아볼 수 있는 정도의 구조화"가 목표이지, 원본 한글 파일의
       모양을 픽셀 단위로 그대로 재현하지 않습니다. 제목(#) 구조는 만들지 않고,
       페이지 레이아웃용으로 쓰인 표는 표가 아니라 줄글로 풀어서 보여줍니다.
    5) 수식 변환은 한컴 공식 수식 명령어를 폭넓게 다루지만 100% 완벽하지는 않습니다.
       모르는 명령어를 만나면 프로그램이 죽거나 내용이 사라지는 게 아니라, 그 부분만
       LaTeX로 안 바뀌고 원래 영어 명령어 글자 그대로 보입니다 (눈에 띄어서 알아채기 쉬움).
       그런 부분을 발견하면 캡처해서 알려주시면 사전에 추가할 수 있습니다.
    6) 중요한 파일은 작업 전에 백업해두는 걸 권장합니다 (혹시 모를 사고 방지 차원).
"""
import argparse
import html
import os
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile


# ============================================================
# 1. 한글 수식(HWP Equation Script) -> LaTeX 변환기
# ============================================================

SPACE_MARKER = '\u0001'  # over/sup/sub 등 구조 파싱이 끝나기 전까지는 진짜 공백 대신 이 마커를 씀
COL_MARKER = '\u0002'    # 행렬/벡터 안에서 열 구분자(겹따옴표)를 표시하는 마커

WORD_MAP = {
    'hbar': r'\hbar',
    'inf': r'\infty', 'infty': r'\infty',
    'union': r'\cup', 'inter': r'\cap', 'smallunion': r'\cup', 'smallinter': r'\cap',
    'times': r'\times', 'nabla': r'\nabla',
    'rightarrow': r'\rightarrow', 'vert': r'\vert', 'simeq': r'\simeq',
    'int': r'\displaystyle\int',
    'oint': r'\displaystyle\oint',
    'dint': r'\displaystyle\iint',
    'tint': r'\displaystyle\iiint',
    'odint': r'\displaystyle\oiint',
    'otint': r'\displaystyle\oiiint',
    'sum': r'\displaystyle\sum', 'lim': r'\lim', 'prod': r'\displaystyle\prod',
    'del': r'\nabla', 'bullet': r'\cdot', 'cdot': r'\cdot',
    'cdots': r'\cdots', 'ldots': r'\ldots',
    'left': r'\left', 'right': r'\right',
    'partial': r'\partial', 'round': r'\partial',  # 한글에서 편미분 기호로 partial과 round 둘 다 씀
    'leq': r'\leq', 'geq': r'\geq', 'neq': r'\neq',
    'triangle': r'\Delta',
    'therefore': r'\therefore',
}
# 그리스 문자: 대소문자에 따라 실제로 다른 기호(소문자 ψ vs 대문자 Ψ)를 뜻하므로
# WORD_MAP과 분리해서, "매칭은 대소문자 무시 + 출력은 실제 첫 글자 대소문자를 따라감" 방식으로 처리
GREEK_LETTERS = [
    'alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta', 'eta', 'theta', 'iota',
    'kappa', 'lambda', 'mu', 'nu', 'xi', 'pi', 'rho', 'sigma', 'tau',
    'upsilon', 'phi', 'chi', 'psi', 'omega',
]

# 글자 꾸밈 명령어(문자 하나를 인자로 받아 장식) - 한컴 공식 목록 기준. 전부 소문자 표준형만 저장하고
# 실제 매칭은 대소문자 구분 없이(HAT/Hat/hat 다 인식) 처리함
ACCENT_FUNCS = {'hat', 'bar', 'vec', 'dot', 'ddot', 'tilde', 'widehat',
                'check', 'acute', 'grave', 'arch', 'dyad', 'under'}
PREFIX_FUNCS = ACCENT_FUNCS | {'sqrt', 'root'}  # 'root'는 sqrt의 동의어
MATRIX_FUNCS = {'pmatrix', 'bmatrix', 'vmatrix', 'dmatrix'}
# 대괄호 없이 세로로 쌓기만 하는 것들 (matrix, cases, pile 계열) - 전부 '#'을 행 구분자로 씀
PLAIN_STACK_FUNCS = {'cases', 'pile', 'lpile', 'rpile'}
# 위/아래첨자를 명시적으로 쓰는 대체 명령어
SUPSUB_FUNCS = {'sup': '^', 'sub': '_'}
# 두 항(앞항 뒷항)을 가져다 조합하는 중위 연산자들
INFIX_OPS = {
    'over': lambda n, d: r'\dfrac{%s}{%s}' % (n, d),
    'atop': lambda n, d: r'{%s \atop %s}' % (n, d),
    'choose': lambda n, d: r'\binom{%s}{%s}' % (n, d),
}


def _match_case(matched_text, canonical_lower):
    """matched_text의 첫 글자가 대문자면 canonical_lower도 첫 글자만 대문자로 바꿔서 반환
    (그리스 문자용: PSI/Psi/psi 전부 인식하되, 소문자로 썼으면 소문자 기호, 대문자로 썼으면 대문자 기호로)"""
    if matched_text[:1].isupper():
        return canonical_lower[:1].upper() + canonical_lower[1:]
    return canonical_lower


def substitute_keywords(s):
    """문자열 전체에서 키워드를 찾아 치환. 대소문자를 구분하지 않고 매칭한다
    (한글 수식 편집기 자체가 대부분 명령어를 대소문자 구분 없이 인식하기 때문).
    단, 그리스 문자는 매칭 후 '실제 입력된 첫 글자의 대소문자'에 따라 소문자(ψ)/대문자(Ψ) 기호를
    다르게 출력한다 — 이 둘은 물리적으로 다른 기호를 뜻하므로.
    그리스 문자와 나머지 키워드를 길이 순으로 하나의 루프에서 같이 처리해야
    'partialPSI'처럼 서로 붙어있는 경우도 순서 꼬임 없이 잡힌다."""
    prefix_names = PREFIX_FUNCS | MATRIX_FUNCS | PLAIN_STACK_FUNCS | set(SUPSUB_FUNCS) | {'matrix', 'not'}
    # 'pi'가 'pile' 앞부분을 먹어버리는 것처럼, 짧은 키워드가 더 긴 구조 명령어 이름과
    # 겹칠 수 있는 경우를 모아서 충돌 시 치환을 건너뜀 (전부 소문자로 비교)
    structural_names = sorted({n.lower() for n in (prefix_names | set(INFIX_OPS))}, key=len, reverse=True)

    # (키워드, 값을 만드는 함수) 쌍을 길이 내림차순으로 하나의 리스트에 합침
    entries = [(name, (lambda matched, name=name: '\\' + _match_case(matched, name))) for name in GREEK_LETTERS]
    entries += [(kw, (lambda matched, v=v: v)) for kw, v in WORD_MAP.items()]
    entries.sort(key=lambda e: len(e[0]), reverse=True)

    for kw, get_value in entries:
        pattern = r'(?<![A-Za-z\\])' + re.escape(kw)

        def repl(m, kw=kw, get_value=get_value):
            full_from_here = m.string[m.start():]
            for sname in structural_names:
                if len(sname) > len(kw) and full_from_here.lower().startswith(sname):
                    return m.group()
            v = get_value(m.group())
            after_str = m.string[m.end():]
            after = after_str[:1]
            if not after.isalpha():
                return v
            # 뒤에 hat/bar/sqrt 같은 접두 함수 이름이 바로 이어지면(예: partialhatQ),
            # 그 함수가 별도 토큰으로 인식되어야 하므로 마커 대신 진짜 공백을 씀
            if any(after_str.lower().startswith(p) for p in prefix_names):
                return v + ' '
            return v + SPACE_MARKER
        s = re.sub(pattern, repl, s, flags=re.IGNORECASE)
    return s


def strip_orphan_close_braces(s):
    """여는 '{' 없이 나온 '}' (원본 오타)를 제거 — \\dfrac{}} 같은 빈 인자 방지"""
    out, depth = [], 0
    for ch in s:
        if ch == '{':
            depth += 1
            out.append(ch)
        elif ch == '}':
            if depth > 0:
                depth -= 1
                out.append(ch)
            # depth==0이면 짝없는 '}'이므로 버림
        else:
            out.append(ch)
    return ''.join(out)


def preprocess(s):
    s = html.unescape(s)
    s = strip_orphan_close_braces(s)
    s = s.replace('\ue04d', '|')  # HWP 수식 폰트(HYhwpEQ) 전용 문자 -> 절댓값 막대
    s = re.sub(r'<=', ' \\\\le ', s)
    s = re.sub(r'>=', ' \\\\ge ', s)
    n_open, n_close = s.count('<'), s.count('>')
    if n_open == n_close and n_open > 0:
        # 짝이 맞음 -> 기댓값/braket 표기 <...> 로 보고 \langle \rangle 로 변환
        s = re.sub(r'<', ' \\\\langle ', s)
        s = re.sub(r'>', ' \\\\rangle ', s)
    # 짝이 안 맞으면(예: "0< x < L") 부등호 비교이므로 그대로 둠
    s = s.replace('~', ' ')  # ~ = 빈칸 한 칸 (한컴 공식 표기)
    s = s.replace('``', COL_MARKER)  # 행렬/벡터 열 구분용 겹따옴표 -> 마커
    s = s.replace('`', ' ')
    s = re.sub(r'(?<=[0-9a-zA-Z}])over\b', ' over', s)
    s = re.sub(r'\bover(?=[a-zA-Z])', 'over ', s)
    s = substitute_keywords(s)
    s = s.replace('+-', ' PLUSMINUS ').replace('-+', ' MINUSPLUS ')
    return s


def tokenize(s):
    tokens = []
    i, n = 0, len(s)
    delims = set('()[]=,')
    while i < n:
        c = s[i]
        if c.isspace() or c == COL_MARKER:
            i += 1
            continue
        if c == '{':
            depth, j = 1, i + 1
            while j < n and depth > 0:
                if s[j] == '{':
                    depth += 1
                elif s[j] == '}':
                    depth -= 1
                j += 1
            tokens.append(s[i:j])
            i = j
        elif c in delims:
            tokens.append(c)
            i += 1
        else:
            j = i
            while j < n and not s[j].isspace() and s[j] not in delims and s[j] != '{' and s[j] != COL_MARKER:
                j += 1
            tokens.append(s[i:j])
            i = j
    return tokens


def find_matching_brace(s):
    """s[0]=='{' 라고 가정, 짝이 맞는 '}' 의 인덱스 반환"""
    depth = 0
    for idx, ch in enumerate(s):
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                return idx
    return len(s)  # 짝이 안 맞으면(원본 오타 등) 끝까지를 내용으로 봄 (마지막 글자 유실 방지)


def convert_word(tok):
    if tok == 'PLUSMINUS':
        return r'\pm'
    if tok == 'MINUSPLUS':
        return r'\mp'
    tok_lower = tok.lower()

    # 행렬 (pmatrix{행1#행2#행3} 형태, '#'이 행 구분자)
    for mfunc in MATRIX_FUNCS:
        if tok_lower.startswith(mfunc) and len(tok) > len(mfunc):
            rest = tok[len(mfunc):]
            if rest.startswith('{'):
                end = find_matching_brace(rest)
                inner, suffix = rest[1:end], rest[end + 1:]
                rows = [convert_expr(r) for r in inner.split('#')]
                body = r' \\ '.join(rows)
                return r'\begin{%s} %s \end{%s}' % (mfunc, body, mfunc) + (convert_word(suffix) if suffix else '')

    # 'matrix{...}' (앞에 p/b/v 없는 것): 겹따옴표(열 구분자)가 있었으면 pmatrix처럼 널찍하게,
    # 없으면 그냥 한 행으로 풀어줌
    if tok_lower.startswith('matrix') and len(tok) > 6:
        rest = tok[6:]
        if rest.startswith('{'):
            end = find_matching_brace(rest)
            inner, suffix = rest[1:end], rest[end + 1:]
            tail = convert_word(suffix) if suffix else ''
            if COL_MARKER in inner:
                parts_raw = inner.split(COL_MARKER)
                first, last = parts_raw[0].strip(), parts_raw[-1].strip()
                # pmatrix가 자체적으로 괄호를 그려주므로, 원본에 이미 있는 감싸는 괄호는 중복이라 벗겨냄
                if first.startswith('(') and last.endswith(')'):
                    parts_raw[0] = parts_raw[0].replace('(', '', 1)
                    parts_raw[-1] = ''.join(parts_raw[-1].rsplit(')', 1))
                cols = [convert_expr(c) for c in parts_raw]
                return r'\begin{pmatrix} %s \end{pmatrix}' % ' & '.join(cols) + tail
            return convert_expr(inner) + tail

    # CASES(경우 나누기), PILE/LPILE/RPILE(세로 쌓기, 괄호 없음) - 전부 '#'이 행 구분자
    for sfunc in PLAIN_STACK_FUNCS:
        if tok_lower.startswith(sfunc) and len(tok) > len(sfunc):
            rest = tok[len(sfunc):]
            if rest.startswith('{'):
                end = find_matching_brace(rest)
                inner, suffix = rest[1:end], rest[end + 1:]
                rows = [convert_expr(r) for r in inner.split('#')]
                env = 'cases' if sfunc == 'cases' else 'matrix'
                body = r' \\ '.join(rows)
                base = r'\begin{%s} %s \end{%s}' % (env, body, env)
                return base + (convert_word(suffix) if suffix else '')

    # SUP/SUB: ^ , _ 를 명시적으로 쓰는 대체 표기
    for sname, sym in SUPSUB_FUNCS.items():
        if tok_lower.startswith(sname) and len(tok) > len(sname):
            rest = tok[len(sname):]
            if rest.startswith('{'):
                end = find_matching_brace(rest)
                inner, suffix = rest[1:end], rest[end + 1:]
                base = '%s{%s}' % (sym, convert_expr(inner))
                return base + (convert_word(suffix) if suffix else '')
            m = re.match(r'^[A-Za-z0-9]+', rest)
            if m:
                target, remainder2 = m.group(), rest[m.end():]
                return '%s{%s}' % (sym, target) + (convert_word(remainder2) if remainder2 else '')

    # NOT: 뒤에 오는 글자/기호에 사선을 긋는 명령 (\not는 중괄호로 감싸지 않고 바로 붙여야 함)
    if tok_lower.startswith('not') and len(tok) > 3:
        rest = tok[3:]
        if rest.startswith('{'):
            end = find_matching_brace(rest)
            inner, suffix = rest[1:end], rest[end + 1:]
            base = r'\not %s' % convert_expr(inner)
            return base + (convert_word(suffix) if suffix else '')
        m = re.match(r'^[A-Za-z0-9=<>]+', rest)
        if m:
            target, remainder2 = m.group(), rest[m.end():]
            return r'\not %s' % target + (convert_word(remainder2) if remainder2 else '')

    for acc in PREFIX_FUNCS:
        if tok_lower.startswith(acc) and len(tok) > len(acc):
            rest = tok[len(acc):]
            cmd = 'sqrt' if acc in ('sqrt', 'root') else acc
            if rest.startswith('{'):
                end = find_matching_brace(rest)
                inner, suffix = rest[1:end], rest[end + 1:]
                base = r'\%s{%s}' % (cmd, convert_expr(inner))
                return base + convert_word(suffix) if suffix else base
            m = re.match(r'^[A-Za-z]+', rest)
            if m:
                target, remainder2 = m.group(), rest[m.end():]
                return r'\%s{%s}' % (cmd, target) + (convert_word(remainder2) if remainder2 else '')
            return r'\%s{%s}' % (cmd, rest) if rest else ('\\' + cmd)
    return tok


def merge_prefix_tokens(tokens):
    """'hat','A' 처럼 따로 떨어진 토큰을 'hatA' 형태로 결합 (대소문자 무시하고 판단)"""
    mergeable = {m.lower() for m in (PREFIX_FUNCS | MATRIX_FUNCS | PLAIN_STACK_FUNCS | set(SUPSUB_FUNCS) | {'matrix', 'not'})}
    out, i = [], 0
    while i < len(tokens):
        t = tokens[i]
        if t.lower() in mergeable and i + 1 < len(tokens):
            out.append(t + tokens[i + 1])
            i += 2
        else:
            out.append(t)
            i += 1
    return out


def merge_sup_sub(tokens):
    """'p', '^2' 처럼 띄어써진 위/아래첨자를 바로 앞 토큰에 붙여서 'p^2'로 합침
    (안 그러면 'over'가 바로 앞 토큰인 '^2'만 분자로 가져가고 'p'를 놓침)"""
    out = []
    for t in tokens:
        if t and t[0] in '^_' and out and out[-1].lower() not in INFIX_OPS:
            out[-1] = out[-1] + t
        else:
            out.append(t)
    return out


def convert_expr(s):
    tokens = merge_sup_sub(merge_prefix_tokens(tokenize(s)))

    def resolve(tok):
        if tok.startswith('{'):
            end = find_matching_brace(tok)
            inner = convert_expr(tok[1:end])
            suffix = tok[end + 1:]
            return '{' + inner + '}' + (convert_word(suffix) if suffix else '')
        return convert_word(tok)

    resolved = [tok if tok.lower() in INFIX_OPS else resolve(tok) for tok in tokens]

    out, i = [], 0
    while i < len(resolved):
        if isinstance(resolved[i], str) and resolved[i].lower() in INFIX_OPS and out and i + 1 < len(resolved):
            num = out.pop()
            den = resolved[i + 1]
            strip = lambda x: x[1:-1] if x.startswith('{') and x.endswith('}') else x
            out.append(INFIX_OPS[resolved[i].lower()](strip(num), strip(den)))
            i += 2
        else:
            out.append(resolved[i])
            i += 1

    return ' '.join(out)


def hwp_eq_to_latex(script):
    result = convert_expr(preprocess(script.strip()))
    return result.replace(SPACE_MARKER, ' ').replace(COL_MARKER, ' ')


# ============================================================
# 2. HWPX (zip + XML) 문서 조립기
# ============================================================

NS = {
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
}


def qn(tag):
    prefix, local = tag.split(':')
    return '{%s}%s' % (NS[prefix], local)


def direct_children(elem, tag):
    """elem의 '직계' 자식 중 해당 태그만 (깊이 상관없이 다 훑는 iter()와 다름)"""
    return [c for c in elem if c.tag == qn(tag)]


def fix_dollar_boundaries(text):
    """닫는 '$' 바로 뒤에 공백 없이 글자/숫자/또 다른 '$'가 붙으면 한 칸 띄워준다."""
    out = []
    dollar_count = 0
    n = len(text)
    for i, ch in enumerate(text):
        out.append(ch)
        if ch == '$':
            dollar_count += 1
            if dollar_count % 2 == 0:
                nxt = text[i + 1] if i + 1 < n else ''
                if nxt.isalnum() or nxt == '$':
                    out.append(' ')
    return ''.join(out)


def get_hwpx_section_roots(hwpx_path):
    """hwpx(zip) 안의 Contents/section*.xml 들을 순서대로 파싱해서 root 리스트 반환"""
    roots = []
    with zipfile.ZipFile(hwpx_path) as z:
        section_names = sorted(
            n for n in z.namelist()
            if re.match(r'Contents/section\d+\.xml$', n)
        )
        for name in section_names:
            with z.open(name) as f:
                roots.append(ET.parse(f).getroot())
    return roots


def hwpx_has_equations(section_roots):
    for root in section_roots:
        if root.find('.//' + qn('hp:equation')) is not None:
            return True
    return False


# --- 수식 있는 문서용: 표 위장 감지 + 수식 변환 포함 전체 파이프라인 ---

def escape_plain_text(text):
    """일반 텍스트(수식 아님) 안의 *, _ 는 마크다운 강조 문법으로 오해될 수 있으므로 이스케이프.
    수식(LaTeX, $...$ 안)에는 적용하면 안 됨 — 거기서는 _가 아래첨자로 실제 의미가 있음."""
    return text.replace('\\', '\\\\').replace('*', '\\*').replace('_', '\\_')


def para_text_full(p_elem, ctx):
    """문단 하나를 문자열로: 텍스트 + 수식(LaTeX) + 중첩 표/이미지 처리"""
    parts = []
    for run in direct_children(p_elem, 'hp:run'):
        for child in run:
            tag = child.tag
            if tag == qn('hp:t'):
                parts.append(escape_plain_text(''.join(child.itertext())))
            elif tag == qn('hp:equation'):
                script_el = child.find(qn('hp:script'))
                if script_el is not None and script_el.text and script_el.text.strip() != 'rm':
                    latex = hwp_eq_to_latex(script_el.text)
                    if latex.strip():
                        parts.append('$%s$' % latex)
            elif tag == qn('hp:tbl'):
                rendered = render_table_full(child, ctx)
                if rendered:
                    parts.append('\n\n' + rendered + '\n\n')
            elif tag == qn('hp:pic'):
                ctx['image_count'] += 1
                parts.append('(그림 %d)' % ctx['image_count'])
            # hp:ctrl (머리말/꼬리말/쪽번호 등 페이지 장식)은 실제 내용이 아니므로 건너뜀
    text = ''.join(parts)
    text = fix_dollar_boundaries(text)
    text = re.sub(r'[ \t]+', ' ', text).strip()
    return text


def cell_paragraphs_text(tc_elem, ctx):
    out = []
    sublist = tc_elem.find(qn('hp:subList'))
    if sublist is None:
        return out
    for p in direct_children(sublist, 'hp:p'):
        t = para_text_full(p, ctx)
        if t:
            out.append(t)
    return out


def render_table_full(tbl_elem, ctx):
    """
    - 한 행에 칸이 1개뿐이면 -> 표가 아니라 페이지 레이아웃용이므로 줄글로 풀어서 반환
    - 칸이 2개 이상인 행만 진짜 표로 렌더링
    """
    output_chunks = []
    pending_rows = []

    def flush():
        if not pending_rows:
            return
        width = max(len(r) for r in pending_rows)
        rows = [r + [''] * (width - len(r)) for r in pending_rows]
        lines = ['| ' + ' | '.join(rows[0]) + ' |',
                 '| ' + ' | '.join(['---'] * width) + ' |']
        for r in rows[1:]:
            lines.append('| ' + ' | '.join(r) + ' |')
        output_chunks.append('\n'.join(lines))
        pending_rows.clear()

    for tr in direct_children(tbl_elem, 'hp:tr'):
        cells = direct_children(tr, 'hp:tc')
        if len(cells) <= 1:
            flush()
            for tc in cells:
                for t in cell_paragraphs_text(tc, ctx):
                    output_chunks.append(t)
        else:
            row = []
            for tc in cells:
                paras = cell_paragraphs_text(tc, ctx)
                row.append(' '.join(paras) if paras else ' ')
            pending_rows.append(row)
    flush()
    return '\n\n'.join(output_chunks)


def convert_hwpx_full(hwpx_path):
    """수식이 있는 hwpx: 표 위장 감지 + 수식 LaTeX 변환 포함 전체 파이프라인"""
    ctx = {'image_count': 0}
    out_chunks = []
    for root in get_hwpx_section_roots(hwpx_path):
        for p in direct_children(root, 'hp:p'):
            t = para_text_full(p, ctx)
            if t:
                out_chunks.append(t)
    return '\n\n'.join(out_chunks)


# --- 수식 없는 문서용: 정말 텍스트만 쭉 뽑기 (표/제목 구조 없음) ---

def convert_hwpx_plain(hwpx_path):
    """수식이 없는 hwpx: 표/제목 등 아무 구조도 만들지 않고, 문단 텍스트만 순서대로 쭉 출력"""
    out_lines = []
    for root in get_hwpx_section_roots(hwpx_path):
        for p in root.iter(qn('hp:p')):
            texts = [''.join(t.itertext()) for t in p.iter(qn('hp:t'))]
            line = escape_plain_text(re.sub(r'[ \t]+', ' ', ''.join(texts)).strip())
            if line:
                out_lines.append(line)
    return '\n\n'.join(out_lines)


# ============================================================
# 3. HWP (구버전 바이너리, OLE2) 처리
# ============================================================

def hwp_has_equations(hwp_path):
    """pyhwp로 EqEdit(수식 컨트롤) 존재 여부 확인. pyhwp 미설치 시 None 반환(모름)."""
    try:
        from hwp5.xmlmodel import Hwp5File
    except ImportError:
        return None
    try:
        f = Hwp5File(hwp_path)
        for section in f.bodytext.sections:
            for model in section.models():
                if model.get('tagname') == 'EqEdit':
                    return True
        return False
    except Exception:
        return None


def convert_hwp_plain(hwp_path):
    """.hwp에서 hwp5txt로 텍스트만 추출 (표/제목 구조 없음)"""
    try:
        result = subprocess.run(
            ['hwp5txt', hwp_path], capture_output=True, text=True, check=True
        )
        return result.stdout.strip()
    except FileNotFoundError:
        raise RuntimeError(
            "hwp5txt 명령을 찾을 수 없습니다. 'pip install pyhwp --break-system-packages'로 설치해주세요."
        )
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"hwp5txt 실행 실패: {e.stderr}")


def hwp_to_hwpx(hwp_path, hwp_app=None):
    """설치된 한글(한컴오피스) 프로그램을 pyhwpx로 조종해서 .hwp -> .hwpx 로 변환.
    hwp_app을 넘기면 그 인스턴스를 재사용(여러 파일 처리 시 매번 새로 켜지 않도록),
    안 넘기면 이 함수 안에서 새로 켰다가 끝나면 닫음."""
    try:
        from pyhwpx import Hwp
    except ImportError:
        raise RuntimeError(
            "pyhwpx가 설치되어 있지 않습니다. 'pip install pyhwpx'로 설치해주세요. "
            "(한글 프로그램이 실제로 설치되어 있어야 동작합니다)"
        )
    owns_app = hwp_app is None
    app = hwp_app or Hwp()
    try:
        hwpx_path = os.path.splitext(hwp_path)[0] + '.hwpx'

        # [미검증 예방 조치] Word COM 자동화에서 경로에 공백이 있으면 파일을 못
        # 찾는 것으로 오작동하는 문제가 확인된 적이 있어서, 한글 자동화도 같은
        # 문제가 있을 가능성에 대비해 open()에 넘기는 경로만 짧은(8.3) 경로로
        # 바꿔서 사용한다. 변환 실패 시(win32api 없거나 8.3 이름 생성이 꺼진 경우)
        # 원래 경로를 그대로 쓴다.
        open_path = hwp_path
        try:
            import win32api
            open_path = win32api.GetShortPathName(hwp_path)
        except Exception:
            pass

        app.open(open_path)
        app.save_as(hwpx_path, format="HWPX")
        # save_as 이후에도 한글 프로그램이 방금 만든 .hwpx를 "열어놓은 채" 붙잡고 있어서,
        # 특히 배치 모드처럼 앱을 재사용할 때는 quit()이 호출되지 않아 잠금이 안 풀림.
        # Clear(1)로 활성 문서만 닫아(변경사항 버림) 잠금을 풀어준다 (앱 자체는 계속 켜둠).
        try:
            app.Clear(1)
        except Exception:
            pass
        return hwpx_path
    finally:
        if owns_app:
            app.quit()
# ============================================================
# 4. 진입점
# ============================================================

def convert(input_path, hwp_app=None):
    ext = os.path.splitext(input_path)[1].lower()
    temp_hwpx_path = None  # .hwp -> .hwpx 로 임시 변환한 경우에만 값이 채워짐 (나중에 삭제용)

    if ext == '.hwp':
        print(f"[{input_path}] .hwp 파일 -> 한글 프로그램으로 .hwpx 변환 중...", file=sys.stderr)
        temp_hwpx_path = hwp_to_hwpx(input_path, hwp_app=hwp_app)
        print(f"  변환 완료(임시): {temp_hwpx_path}", file=sys.stderr)
        input_path = temp_hwpx_path
        ext = '.hwpx'

    try:
        if ext == '.hwpx':
            roots = get_hwpx_section_roots(input_path)
            if hwpx_has_equations(roots):
                print("  수식이 있는 문서로 감지됨 -> 표 위장 감지 + 수식 LaTeX 변환 파이프라인 사용", file=sys.stderr)
                return convert_hwpx_full(input_path)
            else:
                print("  수식이 없는 문서로 감지됨 -> 텍스트만 추출", file=sys.stderr)
                return convert_hwpx_plain(input_path)
        else:
            raise ValueError(f"지원하지 않는 확장자입니다: {ext} (.hwp 또는 .hwpx만 가능)")
    finally:
        # .hwp -> .hwpx 로 만든 임시 파일은 md 변환이 끝나면(성공하든 실패하든) 바로 삭제.
        # 사용자가 직접 고른 .hwpx 원본은 여기 해당 안 되므로 그대로 둠.
        if temp_hwpx_path and os.path.exists(temp_hwpx_path):
            os.remove(temp_hwpx_path)
            print(f"  임시 파일 삭제됨: {temp_hwpx_path}", file=sys.stderr)


def convert_no_hangul(input_path):
    """한글 프로그램 없이(또는 쓰고 싶지 않을 때) 기존 방식대로: .hwp는 텍스트만 추출"""
    ext = os.path.splitext(input_path)[1].lower()
    if ext == '.hwpx':
        roots = get_hwpx_section_roots(input_path)
        if hwpx_has_equations(roots):
            return convert_hwpx_full(input_path)
        return convert_hwpx_plain(input_path)
    elif ext == '.hwp':
        if hwp_has_equations(input_path):
            print(
                "경고: 이 .hwp 파일에 수식이 있지만 한글 프로그램 변환을 쓰지 않아 텍스트만 추출합니다.",
                file=sys.stderr,
            )
        return convert_hwp_plain(input_path)
    else:
        raise ValueError(f"지원하지 않는 확장자입니다: {ext}")


def pick_files_dialog():
    """파일 탐색기 창을 띄워 .hwp/.hwpx 파일을 하나 이상 선택 (Ctrl/Shift로 다중 선택 가능)"""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    paths = filedialog.askopenfilenames(
        title="변환할 HWP/HWPX 파일을 선택하세요 (여러 개 선택 가능: Ctrl/Shift+클릭)",
        filetypes=[("HWP/HWPX 파일", "*.hwp *.hwpx"), ("모든 파일", "*.*")],
    )
    root.destroy()
    return list(paths)


def run_batch(paths):
    """선택된 파일들을 순서대로 변환. .hwp가 하나라도 있으면 한글 프로그램 인스턴스를 하나만 켜서 재사용."""
    hwp_app = None
    needs_hangul = any(os.path.splitext(p)[1].lower() == '.hwp' for p in paths)
    if needs_hangul:
        try:
            from pyhwpx import Hwp
            hwp_app = Hwp()
        except ImportError:
            print(
                "pyhwpx가 없어 .hwp -> .hwpx 자동 변환을 할 수 없습니다. "
                "'pip install pyhwpx' 설치 후 다시 시도해주세요. 일단 텍스트 추출만 진행합니다.",
                file=sys.stderr,
            )

    results = []
    try:
        for path in paths:
            print(f"\n=== 처리 중: {path} ===", file=sys.stderr)
            try:
                if hwp_app is not None:
                    md = convert(path, hwp_app=hwp_app)
                else:
                    md = convert_no_hangul(path)
                out_path = os.path.splitext(path)[0] + '.md'
                with open(out_path, 'w', encoding='utf-8') as f:
                    f.write(md)
                print(f"  -> 완료: {out_path} ({len(md)}자)", file=sys.stderr)
                results.append((path, out_path, None))
            except Exception as e:
                print(f"  -> 실패: {e}", file=sys.stderr)
                results.append((path, None, str(e)))
    finally:
        if hwp_app is not None:
            hwp_app.quit()

    print("\n=== 전체 결과 ===", file=sys.stderr)
    for src, out, err in results:
        status = f"완료 -> {out}" if err is None else f"실패 ({err})"
        print(f"  {os.path.basename(src)}: {status}", file=sys.stderr)
    return results


def main():
    parser = argparse.ArgumentParser(description="HWP/HWPX -> Obsidian용 Markdown 변환기")
    parser.add_argument('input', nargs='?', help='입력 파일 경로 (.hwp 또는 .hwpx). 생략하면 파일 선택 창이 뜸')
    parser.add_argument('-o', '--output', help='출력 마크다운 파일 경로 (input을 직접 지정했을 때만)')
    parser.add_argument('--no-hangul', action='store_true',
                         help='한글 프로그램을 쓰지 않음 (.hwp 수식은 변환 안 되고 텍스트만 추출됨)')
    args = parser.parse_args()

    if args.input:
        # 기존처럼 파일 하나 직접 지정하는 방식 (스크립트 등에서 자동화할 때 사용)
        if args.no_hangul:
            md = convert_no_hangul(args.input)
        else:
            md = convert(args.input)
        out_path = args.output or (os.path.splitext(args.input)[0] + '.md')
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(md)
        print(f"완료: {out_path} ({len(md)}자)", file=sys.stderr)
    else:
        # 인자 없이 실행하면 파일 선택 창 띄우기
        paths = pick_files_dialog()
        if not paths:
            print("선택된 파일이 없습니다.", file=sys.stderr)
            return
        run_batch(paths)


if __name__ == '__main__':
    main()
