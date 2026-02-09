"""Jupyter Notebook (.ipynb) 생성 모듈.

nbformat을 사용하여 문항과 풀이를 노트북으로 구성한다.
"""

from pathlib import Path
from dataclasses import dataclass

import nbformat
from nbformat.v4 import new_notebook, new_markdown_cell, new_code_cell


@dataclass
class SolvedQuestion:
    """풀이가 포함된 문항."""
    number: int
    title: str
    body: str
    code: str


# 공통 import 코드
COMMON_IMPORTS = """\
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats

plt.rcParams['font.family'] = 'Malgun Gothic'  # 한글 폰트
plt.rcParams['axes.unicode_minus'] = False      # 마이너스 부호
"""


def build_notebook(
    solved_questions: list[SolvedQuestion],
    title: str = "데이터분석 실기 풀이",
    output_path: str | None = None,
) -> nbformat.NotebookNode:
    """문항과 풀이를 Jupyter Notebook으로 생성한다.

    구조: 제목 셀 → 공통 import 셀 → (문항 마크다운 + 풀이 코드) × N

    Args:
        solved_questions: SolvedQuestion 리스트
        title: 노트북 제목
        output_path: 저장할 파일 경로 (None이면 저장하지 않음)

    Returns:
        nbformat.NotebookNode 객체
    """
    nb = new_notebook()

    # 커널 정보 설정
    nb.metadata.kernelspec = {
        "display_name": "Python 3",
        "language": "python",
        "name": "python3",
    }
    nb.metadata.language_info = {
        "name": "python",
        "version": "3.10.0",
    }

    # 제목 셀
    nb.cells.append(new_markdown_cell(f"# {title}"))

    # 공통 import 셀
    nb.cells.append(new_code_cell(COMMON_IMPORTS))

    # 문항별 셀
    for q in solved_questions:
        # 문항 마크다운 셀
        md_content = f"## 문제 {q.number}\n\n{q.body}"
        nb.cells.append(new_markdown_cell(md_content))

        # 풀이 코드 셀
        if q.code:
            nb.cells.append(new_code_cell(q.code))

    # 파일 저장
    if output_path:
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            nbformat.write(nb, f)

    return nb
