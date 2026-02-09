"""Word 문서(.docx)에서 텍스트와 수식을 추출한다.

python-docx의 paragraph.text는 수식(<m:oMath>)을 누락하므로,
lxml을 사용하여 <w:p> 요소를 직접 순회한다.
"""

from pathlib import Path
from lxml import etree
from docx import Document

from .omml_parser import omml_to_latex

# 네임스페이스
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

W_P = f"{{{WORD_NS}}}p"       # <w:p> 문단
W_R = f"{{{WORD_NS}}}r"       # <w:r> 런
W_T = f"{{{WORD_NS}}}t"       # <w:t> 텍스트
W_TBL = f"{{{WORD_NS}}}tbl"   # <w:tbl> 테이블
W_TR = f"{{{WORD_NS}}}tr"     # <w:tr> 테이블 행
W_TC = f"{{{WORD_NS}}}tc"     # <w:tc> 테이블 셀
W_BR = f"{{{WORD_NS}}}br"     # <w:br> 줄바꿈
W_TAB = f"{{{WORD_NS}}}tab"   # <w:tab> 탭
M_OMATH = f"{{{MATH_NS}}}oMath"        # 인라인 수식
M_OMATHPARA = f"{{{MATH_NS}}}oMathPara" # 디스플레이 수식


def _extract_run_text(run_elem) -> str:
    """<w:r> 요소에서 텍스트를 추출한다."""
    parts = []
    for child in run_elem:
        if child.tag == W_T:
            parts.append(child.text or "")
        elif child.tag == W_BR:
            parts.append("\n")
        elif child.tag == W_TAB:
            parts.append("\t")
    return "".join(parts)


def _process_paragraph(para_elem) -> str:
    """<w:p> 요소를 처리하여 텍스트+LaTeX 수식 문자열을 반환한다."""
    parts = []

    for child in para_elem:
        tag = child.tag

        if tag == W_R:
            # 일반 텍스트 런
            parts.append(_extract_run_text(child))

        elif tag == M_OMATH:
            # 인라인 수식
            latex = omml_to_latex(child)
            if latex:
                parts.append(f" ${latex}$ ")

        elif tag == M_OMATHPARA:
            # 디스플레이 수식 (블록 수식)
            # m:oMathPara 안에 m:oMath가 있음
            omath_elems = child.findall(f".//{M_OMATH}")
            for omath in omath_elems:
                latex = omml_to_latex(omath)
                if latex:
                    parts.append(f"\n$${latex}$$\n")

    return "".join(parts)


def _process_table(tbl_elem) -> str:
    """<w:tbl> 요소를 처리하여 마크다운 테이블 문자열을 반환한다."""
    rows = []
    for tr in tbl_elem:
        if tr.tag != W_TR:
            continue
        cells = []
        for tc in tr:
            if tc.tag != W_TC:
                continue
            # 셀 안의 모든 문단 처리
            cell_parts = []
            for para in tc:
                if para.tag == W_P:
                    text = _process_paragraph(para).strip()
                    if text:
                        cell_parts.append(text)
            cells.append(" ".join(cell_parts))
        if cells:
            rows.append(cells)

    if not rows:
        return ""

    # 마크다운 테이블 생성
    lines = []
    # 헤더
    lines.append("| " + " | ".join(rows[0]) + " |")
    lines.append("| " + " | ".join(["---"] * len(rows[0])) + " |")
    # 데이터 행
    for row in rows[1:]:
        # 열 수가 부족하면 빈 셀 추가
        while len(row) < len(rows[0]):
            row.append("")
        lines.append("| " + " | ".join(row[:len(rows[0])]) + " |")

    return "\n".join(lines)


def extract_document(docx_path: str) -> str:
    """Word 문서에서 전체 텍스트를 추출한다.

    수식은 LaTeX로 변환되어 $...$ 또는 $$...$$ 안에 포함된다.
    테이블은 마크다운 테이블로 변환된다.

    Args:
        docx_path: .docx 파일 경로

    Returns:
        추출된 전체 텍스트 (수식 포함)
    """
    path = Path(docx_path)
    if not path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {docx_path}")
    if not path.suffix.lower() == ".docx":
        raise ValueError(f"지원하지 않는 파일 형식입니다: {path.suffix}")

    doc = Document(str(path))
    body = doc.element.body

    paragraphs = []

    for elem in body:
        tag = elem.tag

        if tag == W_P:
            text = _process_paragraph(elem)
            paragraphs.append(text)

        elif tag == W_TBL:
            table_md = _process_table(elem)
            if table_md:
                paragraphs.append("\n" + table_md + "\n")

    # 연속된 빈 줄을 최대 2줄로 제한
    result = "\n".join(paragraphs)
    while "\n\n\n" in result:
        result = result.replace("\n\n\n", "\n\n")

    return result.strip()
