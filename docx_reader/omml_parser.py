"""OMML (Office Math Markup Language) XML → LaTeX 변환기.

Word 2016 수식 편집기로 작성된 수식(<m:oMath>)을 LaTeX 문자열로 변환한다.
python-docx의 paragraph.text는 수식을 누락하므로 lxml로 직접 XML을 파싱한다.
"""

from lxml import etree

# OMML 네임스페이스
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

NSMAP = {
    "m": MATH_NS,
    "w": WORD_NS,
}

# 유니코드 → LaTeX 심볼 매핑
UNICODE_TO_LATEX = {
    # 그리스 소문자
    "\u03b1": r"\alpha",
    "\u03b2": r"\beta",
    "\u03b3": r"\gamma",
    "\u03b4": r"\delta",
    "\u03b5": r"\epsilon",
    "\u03b6": r"\zeta",
    "\u03b7": r"\eta",
    "\u03b8": r"\theta",
    "\u03b9": r"\iota",
    "\u03ba": r"\kappa",
    "\u03bb": r"\lambda",
    "\u03bc": r"\mu",
    "\u03bd": r"\nu",
    "\u03be": r"\xi",
    "\u03c0": r"\pi",
    "\u03c1": r"\rho",
    "\u03c3": r"\sigma",
    "\u03c4": r"\tau",
    "\u03c5": r"\upsilon",
    "\u03c6": r"\phi",
    "\u03c7": r"\chi",
    "\u03c8": r"\psi",
    "\u03c9": r"\omega",
    # 그리스 대문자
    "\u0393": r"\Gamma",
    "\u0394": r"\Delta",
    "\u0398": r"\Theta",
    "\u039b": r"\Lambda",
    "\u039e": r"\Xi",
    "\u03a0": r"\Pi",
    "\u03a3": r"\Sigma",
    "\u03a6": r"\Phi",
    "\u03a8": r"\Psi",
    "\u03a9": r"\Omega",
    # 수학 기호
    "\u00b1": r"\pm",
    "\u00d7": r"\times",
    "\u00f7": r"\div",
    "\u2202": r"\partial",
    "\u2207": r"\nabla",
    "\u221a": r"\sqrt",
    "\u221e": r"\infty",
    "\u2211": r"\sum",
    "\u220f": r"\prod",
    "\u222b": r"\int",
    "\u222c": r"\iint",
    "\u222d": r"\iiint",
    "\u2260": r"\neq",
    "\u2264": r"\leq",
    "\u2265": r"\geq",
    "\u2248": r"\approx",
    "\u2261": r"\equiv",
    "\u2208": r"\in",
    "\u2209": r"\notin",
    "\u2282": r"\subset",
    "\u2283": r"\supset",
    "\u2286": r"\subseteq",
    "\u2287": r"\supseteq",
    "\u222a": r"\cup",
    "\u2229": r"\cap",
    "\u2205": r"\emptyset",
    "\u2200": r"\forall",
    "\u2203": r"\exists",
    "\u00ac": r"\neg",
    "\u2227": r"\wedge",
    "\u2228": r"\vee",
    "\u2192": r"\rightarrow",
    "\u2190": r"\leftarrow",
    "\u21d2": r"\Rightarrow",
    "\u21d0": r"\Leftarrow",
    "\u21d4": r"\Leftrightarrow",
    "\u22c5": r"\cdot",
    "\u2026": r"\ldots",
    "\u22ef": r"\cdots",
    "\u22ee": r"\vdots",
    "\u22f1": r"\ddots",
    # 위첨자/아래첨자 숫자 (문서에서 종종 등장)
    "\u2070": "^{0}",
    "\u00b9": "^{1}",
    "\u00b2": "^{2}",
    "\u00b3": "^{3}",
    "\u2074": "^{4}",
    "\u2075": "^{5}",
    "\u2076": "^{6}",
    "\u2077": "^{7}",
    "\u2078": "^{8}",
    "\u2079": "^{9}",
    # 악센트 관련
    "\u0302": r"\hat",
    "\u0304": r"\bar",
    "\u0307": r"\dot",
    "\u0308": r"\ddot",
    "\u0303": r"\tilde",
    "\u20d7": r"\vec",
}

# Nary 연산자 매핑 (유니코드 문자 → LaTeX 명령)
NARY_MAP = {
    "\u2211": r"\sum",
    "\u220f": r"\prod",
    "\u222b": r"\int",
    "\u222c": r"\iint",
    "\u222d": r"\iiint",
    "\u222e": r"\oint",
}

# 악센트 매핑
ACCENT_MAP = {
    "\u0302": r"\hat",
    "\u0304": r"\bar",
    "\u0307": r"\dot",
    "\u0308": r"\ddot",
    "\u0303": r"\tilde",
    "\u20d7": r"\vec",
    "\u0305": r"\overline",
    "\u23de": r"\overbrace",
    "\u23df": r"\underbrace",
}


def _qn(tag: str) -> str:
    """네임스페이스 약칭을 전체 URI로 확장한다. 예: 'm:f' → '{uri}f'."""
    prefix, local = tag.split(":")
    return f"{{{NSMAP[prefix]}}}{local}"


def _find(elem, tag: str):
    """elem 하위에서 tag에 해당하는 첫 번째 자식을 찾는다."""
    return elem.find(_qn(tag))


def _findall(elem, tag: str):
    """elem 하위에서 tag에 해당하는 모든 자식을 찾는다."""
    return elem.findall(_qn(tag))


def _get_val(elem, tag: str) -> str:
    """m:xxxPr 내의 m:val 속성 값을 가져온다."""
    child = _find(elem, tag)
    if child is not None:
        return child.get(_qn("m:val"), "")
    return ""


def _escape_latex(text: str) -> str:
    """유니코드 문자를 LaTeX 심볼로 변환한다."""
    result = []
    for ch in text:
        if ch in UNICODE_TO_LATEX:
            latex = UNICODE_TO_LATEX[ch]
            # 명령어 뒤에 공백 추가 (backslash로 시작하는 경우)
            if latex.startswith("\\") and not latex.startswith("^"):
                result.append(latex + " ")
            else:
                result.append(latex)
        else:
            result.append(ch)
    return "".join(result)


class OMMLToLatex:
    """OMML XML 요소를 LaTeX 문자열로 변환하는 클래스.

    사용법:
        converter = OMMLToLatex()
        latex_str = converter.convert(omath_element)
    """

    def __init__(self):
        # 태그별 핸들러 매핑
        self._handlers = {
            _qn("m:f"): self._handle_frac,
            _qn("m:sSup"): self._handle_ssup,
            _qn("m:sSub"): self._handle_ssub,
            _qn("m:sSubSup"): self._handle_ssubsup,
            _qn("m:rad"): self._handle_rad,
            _qn("m:nary"): self._handle_nary,
            _qn("m:d"): self._handle_delim,
            _qn("m:func"): self._handle_func,
            _qn("m:acc"): self._handle_acc,
            _qn("m:bar"): self._handle_bar,
            _qn("m:m"): self._handle_matrix,
            _qn("m:eqArr"): self._handle_eqarr,
            _qn("m:limLow"): self._handle_limlow,
            _qn("m:limUpp"): self._handle_limupp,
            _qn("m:groupChr"): self._handle_groupchr,
            _qn("m:borderBox"): self._handle_borderbox,
            _qn("m:box"): self._handle_box,
            _qn("m:sPre"): self._handle_spre,
            _qn("m:r"): self._handle_run,
        }

    def convert(self, elem) -> str:
        """OMML 요소 트리를 LaTeX 문자열로 변환한다."""
        if elem is None:
            return ""
        return self._process(elem).strip()

    def _process(self, elem) -> str:
        """요소를 재귀적으로 처리한다."""
        tag = elem.tag

        # 등록된 핸들러가 있으면 사용
        handler = self._handlers.get(tag)
        if handler:
            return handler(elem)

        # 핸들러가 없으면 자식들을 순서대로 처리
        parts = []
        for child in elem:
            parts.append(self._process(child))
        return "".join(parts)

    def _process_children(self, elem) -> str:
        """elem의 모든 자식을 처리하여 결합한다."""
        parts = []
        for child in elem:
            parts.append(self._process(child))
        return "".join(parts)

    def _get_element_text(self, elem) -> str:
        """m:e (요소) 내용을 처리한다."""
        e = _find(elem, "m:e")
        if e is not None:
            return self._process_children(e)
        return ""

    # ── 핸들러들 ──────────────────────────────────────────────

    def _handle_run(self, elem) -> str:
        """m:r (수식 런) - 텍스트 추출."""
        t = _find(elem, "m:t")
        if t is not None and t.text:
            return _escape_latex(t.text)
        # w:t 안에 있을 수도 있음
        wt = _find(elem, "w:t")
        if wt is not None and wt.text:
            return _escape_latex(wt.text)
        return ""

    def _handle_frac(self, elem) -> str:
        """m:f (분수) → \\frac{num}{den}."""
        # 분수 유형 확인
        fpr = _find(elem, "m:fPr")
        frac_type = ""
        if fpr is not None:
            frac_type = _get_val(fpr, "m:type")

        num_elem = _find(elem, "m:num")
        den_elem = _find(elem, "m:den")
        num = self._process_children(num_elem) if num_elem is not None else ""
        den = self._process_children(den_elem) if den_elem is not None else ""

        if frac_type == "lin":
            # 인라인 분수: a/b
            return f"{num}/{den}"
        return rf"\frac{{{num}}}{{{den}}}"

    def _handle_ssup(self, elem) -> str:
        """m:sSup (위첨자) → base^{sup}."""
        e = _find(elem, "m:e")
        sup = _find(elem, "m:sup")
        base = self._process_children(e) if e is not None else ""
        sup_text = self._process_children(sup) if sup is not None else ""
        return f"{base}^{{{sup_text}}}"

    def _handle_ssub(self, elem) -> str:
        """m:sSub (아래첨자) → base_{sub}."""
        e = _find(elem, "m:e")
        sub = _find(elem, "m:sub")
        base = self._process_children(e) if e is not None else ""
        sub_text = self._process_children(sub) if sub is not None else ""
        return f"{base}_{{{sub_text}}}"

    def _handle_ssubsup(self, elem) -> str:
        """m:sSubSup (아래+위첨자) → base_{sub}^{sup}."""
        e = _find(elem, "m:e")
        sub = _find(elem, "m:sub")
        sup = _find(elem, "m:sup")
        base = self._process_children(e) if e is not None else ""
        sub_text = self._process_children(sub) if sub is not None else ""
        sup_text = self._process_children(sup) if sup is not None else ""
        return f"{base}_{{{sub_text}}}^{{{sup_text}}}"

    def _handle_rad(self, elem) -> str:
        """m:rad (근호) → \\sqrt{e} 또는 \\sqrt[deg]{e}."""
        # 차수 확인
        rad_pr = _find(elem, "m:radPr")
        deg_hide = False
        if rad_pr is not None:
            dh = _find(rad_pr, "m:degHide")
            if dh is not None:
                deg_hide = dh.get(_qn("m:val"), "") == "1"

        deg = _find(elem, "m:deg")
        e = _find(elem, "m:e")
        content = self._process_children(e) if e is not None else ""

        if deg is not None and not deg_hide:
            deg_text = self._process_children(deg).strip()
            if deg_text:
                return rf"\sqrt[{deg_text}]{{{content}}}"

        return rf"\sqrt{{{content}}}"

    def _handle_nary(self, elem) -> str:
        """m:nary (N-항 연산자: 합, 적분 등) → \\sum_{sub}^{sup} e."""
        nary_pr = _find(elem, "m:naryPr")

        # 연산자 문자 확인
        op_char = "\u2211"  # 기본값: 합
        if nary_pr is not None:
            chr_elem = _find(nary_pr, "m:chr")
            if chr_elem is not None:
                op_char = chr_elem.get(_qn("m:val"), op_char)

        latex_op = NARY_MAP.get(op_char, r"\sum")

        # 상한/하한 위치 확인
        lim_loc = ""
        if nary_pr is not None:
            lim_loc = _get_val(nary_pr, "m:limLoc")

        sub = _find(elem, "m:sub")
        sup = _find(elem, "m:sup")
        e = _find(elem, "m:e")

        sub_text = self._process_children(sub).strip() if sub is not None else ""
        sup_text = self._process_children(sup).strip() if sup is not None else ""
        content = self._process_children(e) if e is not None else ""

        result = latex_op
        if sub_text:
            result += f"_{{{sub_text}}}"
        if sup_text:
            result += f"^{{{sup_text}}}"
        result += f" {content}"

        return result

    def _handle_delim(self, elem) -> str:
        """m:d (구분자/괄호) → \\left( ... \\right)."""
        d_pr = _find(elem, "m:dPr")
        beg_chr = "("
        end_chr = ")"
        sep_chr = "|"

        if d_pr is not None:
            bc = _find(d_pr, "m:begChr")
            if bc is not None:
                beg_chr = bc.get(_qn("m:val"), "(")
            ec = _find(d_pr, "m:endChr")
            if ec is not None:
                end_chr = ec.get(_qn("m:val"), ")")
            sc = _find(d_pr, "m:sepChr")
            if sc is not None:
                sep_chr = sc.get(_qn("m:val"), "|")

        # 특수 괄호 매핑
        brace_map = {
            "{": r"\{",
            "}": r"\}",
            "": ".",  # 빈 구분자
            "|": "|",
            "‖": r"\|",
            "⌈": r"\lceil",
            "⌉": r"\rceil",
            "⌊": r"\lfloor",
            "⌋": r"\rfloor",
            "⟨": r"\langle",
            "⟩": r"\rangle",
        }

        left = brace_map.get(beg_chr, beg_chr)
        right = brace_map.get(end_chr, end_chr)

        # m:e 요소들 처리
        elements = _findall(elem, "m:e")
        parts = []
        for e in elements:
            parts.append(self._process_children(e))

        sep = f" {sep_chr} " if len(parts) > 1 else ""
        inner = sep.join(parts)

        return rf"\left{left} {inner} \right{right}"

    def _handle_func(self, elem) -> str:
        """m:func (함수) → \\funcname{arg}."""
        fname_elem = _find(elem, "m:fName")
        e = _find(elem, "m:e")

        fname = ""
        if fname_elem is not None:
            fname = self._process_children(fname_elem).strip()

        content = self._process_children(e) if e is not None else ""

        # 이미 LaTeX 명령이면 그대로 사용
        known_funcs = {
            "sin", "cos", "tan", "cot", "sec", "csc",
            "arcsin", "arccos", "arctan",
            "sinh", "cosh", "tanh",
            "log", "ln", "exp", "lim", "max", "min",
            "sup", "inf", "det", "dim", "gcd",
        }

        # 함수명에서 백슬래시 제거 후 확인
        clean_name = fname.replace("\\", "").strip()
        if clean_name in known_funcs:
            return rf"\{clean_name} {content}"

        return f"{fname} {content}"

    def _handle_acc(self, elem) -> str:
        """m:acc (악센트) → \\hat{e}, \\bar{e} 등."""
        acc_pr = _find(elem, "m:accPr")
        accent_char = "\u0302"  # 기본: hat
        if acc_pr is not None:
            chr_elem = _find(acc_pr, "m:chr")
            if chr_elem is not None:
                accent_char = chr_elem.get(_qn("m:val"), accent_char)

        e = _find(elem, "m:e")
        content = self._process_children(e) if e is not None else ""

        latex_accent = ACCENT_MAP.get(accent_char, r"\hat")
        return f"{latex_accent}{{{content}}}"

    def _handle_bar(self, elem) -> str:
        """m:bar (윗줄/밑줄) → \\overline{e} 또는 \\underline{e}."""
        bar_pr = _find(elem, "m:barPr")
        pos = "top"
        if bar_pr is not None:
            pos_elem = _find(bar_pr, "m:pos")
            if pos_elem is not None:
                pos = pos_elem.get(_qn("m:val"), "top")

        e = _find(elem, "m:e")
        content = self._process_children(e) if e is not None else ""

        if pos == "bot":
            return rf"\underline{{{content}}}"
        return rf"\overline{{{content}}}"

    def _handle_matrix(self, elem) -> str:
        """m:m (행렬) → \\begin{matrix} ... \\end{matrix}."""
        rows = _findall(elem, "m:mr")
        row_strs = []
        for row in rows:
            cells = _findall(row, "m:e")
            cell_strs = []
            for cell in cells:
                cell_strs.append(self._process_children(cell))
            row_strs.append(" & ".join(cell_strs))

        inner = r" \\ ".join(row_strs)
        return rf"\begin{{matrix}} {inner} \end{{matrix}}"

    def _handle_eqarr(self, elem) -> str:
        """m:eqArr (수식 배열) → \\begin{aligned} ... \\end{aligned}."""
        equations = _findall(elem, "m:e")
        parts = []
        for eq in equations:
            parts.append(self._process_children(eq))

        inner = r" \\ ".join(parts)
        return rf"\begin{{aligned}} {inner} \end{{aligned}}"

    def _handle_limlow(self, elem) -> str:
        """m:limLow (하한) → base_{lim}."""
        e = _find(elem, "m:e")
        lim = _find(elem, "m:lim")
        base = self._process_children(e) if e is not None else ""
        lim_text = self._process_children(lim) if lim is not None else ""
        return f"{base}_{{{lim_text}}}"

    def _handle_limupp(self, elem) -> str:
        """m:limUpp (상한) → base^{lim}."""
        e = _find(elem, "m:e")
        lim = _find(elem, "m:lim")
        base = self._process_children(e) if e is not None else ""
        lim_text = self._process_children(lim) if lim is not None else ""
        return f"{base}^{{{lim_text}}}"

    def _handle_groupchr(self, elem) -> str:
        """m:groupChr (그룹 문자) → \\overbrace{e} 등."""
        gpr = _find(elem, "m:groupChrPr")
        chr_val = "\u23df"  # 기본: underbrace
        pos = "bot"

        if gpr is not None:
            chr_elem = _find(gpr, "m:chr")
            if chr_elem is not None:
                chr_val = chr_elem.get(_qn("m:val"), chr_val)
            pos_elem = _find(gpr, "m:pos")
            if pos_elem is not None:
                pos = pos_elem.get(_qn("m:val"), "bot")

        e = _find(elem, "m:e")
        content = self._process_children(e) if e is not None else ""

        if chr_val == "\u23de" or pos == "top":
            return rf"\overbrace{{{content}}}"
        return rf"\underbrace{{{content}}}"

    def _handle_borderbox(self, elem) -> str:
        """m:borderBox (테두리 박스) → \\boxed{e}."""
        e = _find(elem, "m:e")
        content = self._process_children(e) if e is not None else ""
        return rf"\boxed{{{content}}}"

    def _handle_box(self, elem) -> str:
        """m:box (박스) → 내용만 추출."""
        e = _find(elem, "m:e")
        return self._process_children(e) if e is not None else ""

    def _handle_spre(self, elem) -> str:
        """m:sPre (전치 첨자) → {}_{sub}^{sup} base."""
        sub = _find(elem, "m:sub")
        sup = _find(elem, "m:sup")
        e = _find(elem, "m:e")

        sub_text = self._process_children(sub) if sub is not None else ""
        sup_text = self._process_children(sup) if sup is not None else ""
        content = self._process_children(e) if e is not None else ""

        pre = "{}^"
        result = ""
        if sub_text:
            result += f"{{}}_{{{sub_text}}}"
        if sup_text:
            result += f"{{}}^{{{sup_text}}}"
        result += f" {content}"

        return result


# 모듈 레벨 싱글톤
_converter = OMMLToLatex()


def omml_to_latex(omath_elem) -> str:
    """OMML <m:oMath> 요소를 LaTeX 문자열로 변환한다.

    Args:
        omath_elem: lxml Element (<m:oMath> 또는 <m:oMathPara>)

    Returns:
        LaTeX 문자열
    """
    return _converter.convert(omath_elem)
