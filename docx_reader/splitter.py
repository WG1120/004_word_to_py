"""추출된 문서 텍스트를 개별 문항으로 분리한다.

다양한 문항 번호 패턴을 지원:
  - "1.", "2.", "3." (숫자+마침표)
  - "1)", "2)", "3)" (숫자+괄호)
  - "문제 1", "문제 2" (한글 접두어)
  - "[1]", "[2]" (대괄호)
  - "Q1", "Q2" (영문 접두어)
  - "제1문", "제2문" (한자식)
"""

import re
from dataclasses import dataclass


@dataclass
class Question:
    """분리된 문항."""
    number: int       # 문항 번호 (1부터 시작)
    title: str        # 문항 제목 (번호 포함)
    body: str         # 문항 본문 (제목 포함 전체 텍스트)


# 문항 시작을 감지하는 정규식 패턴들 (우선순위 순)
_QUESTION_PATTERNS = [
    # "문제 1", "문제 1.", "문제1" 등
    re.compile(r"^문제\s*(\d+)[.)]?\s", re.MULTILINE),
    # "제1문", "제2문" 등
    re.compile(r"^제\s*(\d+)\s*문", re.MULTILINE),
    # "[1]", "[2]" 등 (줄 시작)
    re.compile(r"^\[(\d+)\]", re.MULTILINE),
    # "Q1.", "Q1)", "Q1 " 등
    re.compile(r"^Q(\d+)[.):\s]", re.MULTILINE),
    # "1.", "2.", "3." (줄 시작, 단 소수점과 구분하기 위해 뒤에 공백 필요)
    re.compile(r"^(\d+)\.\s", re.MULTILINE),
    # "1)", "2)", "3)" (줄 시작)
    re.compile(r"^(\d+)\)\s", re.MULTILINE),
]


def split_questions(text: str) -> list[Question]:
    """텍스트를 문항 단위로 분리한다.

    여러 문항 번호 패턴을 시도하고, 가장 많은 매치를 반환하는 패턴을 선택한다.
    패턴이 전혀 매치되지 않으면 전체 텍스트를 하나의 문항으로 반환한다.

    Args:
        text: 추출된 문서 텍스트

    Returns:
        Question 객체 리스트
    """
    if not text.strip():
        return []

    best_matches = []
    best_pattern = None

    for pattern in _QUESTION_PATTERNS:
        matches = list(pattern.finditer(text))
        if len(matches) > len(best_matches):
            best_matches = matches
            best_pattern = pattern

    # 매치가 없으면 전체를 하나의 문항으로
    if not best_matches:
        return [Question(number=1, title="문제 1", body=text.strip())]

    questions = []
    for i, match in enumerate(best_matches):
        number = int(match.group(1))
        start = match.start()

        # 다음 문항 시작 또는 문서 끝까지
        if i + 1 < len(best_matches):
            end = best_matches[i + 1].start()
        else:
            end = len(text)

        body = text[start:end].strip()

        # 제목: 첫 번째 줄 (또는 첫 100자)
        first_line = body.split("\n")[0].strip()
        title = first_line[:100] if len(first_line) > 100 else first_line

        questions.append(Question(number=number, title=title, body=body))

    return questions
