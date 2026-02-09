"""OpenAI API 클라이언트 래퍼.

재시도 로직과 응답 후처리를 포함한다.
"""

import re
import time

from openai import OpenAI, RateLimitError, APIError

import config
from .prompts import SYSTEM_PROMPT, build_user_prompt


def _strip_code_fences(text: str) -> str:
    """마크다운 코드 펜스를 제거한다.

    GPT가 ```python ... ``` 로 감싸서 응답하는 경우를 처리한다.
    """
    # ```python\n...\n``` 또는 ```\n...\n``` 패턴 제거
    pattern = r"^```(?:python|py)?\s*\n(.*?)```\s*$"
    match = re.match(pattern, text.strip(), re.DOTALL)
    if match:
        return match.group(1).strip()
    return text.strip()


def solve_question(
    question_body: str,
    model: str | None = None,
    api_key: str | None = None,
) -> str:
    """GPT API를 호출하여 문제의 파이썬 풀이를 생성한다.

    Args:
        question_body: 문제 본문 (LaTeX 수식 포함 가능)
        model: 사용할 모델명 (기본: config.DEFAULT_MODEL)
        api_key: OpenAI API 키 (기본: config.OPENAI_API_KEY)

    Returns:
        파이썬 코드 문자열

    Raises:
        RuntimeError: 최대 재시도 횟수 초과 시
    """
    model = model or config.DEFAULT_MODEL
    api_key = api_key or config.OPENAI_API_KEY

    if not api_key:
        raise RuntimeError(
            "OPENAI_API_KEY가 설정되지 않았습니다. "
            ".env 파일에 OPENAI_API_KEY를 설정해주세요."
        )

    client = OpenAI(api_key=api_key)
    user_prompt = build_user_prompt(question_body)

    last_error = None
    for attempt in range(config.MAX_RETRIES):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=config.TEMPERATURE,
            )
            content = response.choices[0].message.content or ""
            return _strip_code_fences(content)

        except RateLimitError as e:
            last_error = e
            delay = config.RETRY_BASE_DELAY * (2 ** attempt)
            print(f"  Rate limit 초과. {delay}초 후 재시도... ({attempt + 1}/{config.MAX_RETRIES})")
            time.sleep(delay)

        except APIError as e:
            last_error = e
            delay = config.RETRY_BASE_DELAY * (2 ** attempt)
            print(f"  API 오류: {e}. {delay}초 후 재시도... ({attempt + 1}/{config.MAX_RETRIES})")
            time.sleep(delay)

    raise RuntimeError(
        f"API 호출이 {config.MAX_RETRIES}회 실패했습니다: {last_error}"
    )
