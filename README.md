# Word to Jupyter Notebook Solver

MS Word 2016 파일(.docx)에 포함된 데이터분석 실기 문항을 읽어, ChatGPT API(gpt-4o-mini)로 파이썬 풀이를 생성하고, Jupyter Notebook(.ipynb)으로 출력하는 CLI 프로그램입니다.

## 주요 기능

- **OMML 수식 추출**: Word 수식 편집기(OMML)로 작성된 수식을 LaTeX로 정확히 변환
  - 분수(`\frac`), 첨자, 근호(`\sqrt`), 합/적분(`\sum`, `\int`), 행렬, 괄호 등 약 20개 수식 요소 지원
  - 그리스 문자, 수학기호 등 유니코드 → LaTeX 심볼 자동 매핑
- **문항 자동 분리**: `1.`, `1)`, `문제 1`, `[1]`, `Q1`, `제1문` 등 다양한 번호 패턴 인식
- **GPT 풀이 생성**: numpy, pandas, scipy, matplotlib 활용한 파이썬 코드 자동 생성
- **Jupyter Notebook 출력**: 문항 마크다운 + 풀이 코드가 포함된 .ipynb 파일 생성

## 프로젝트 구조

```
004_word_to_py/
├── main.py                  # CLI 진입점
├── config.py                # 설정 (모델명, 환경변수 로드)
├── requirements.txt         # 의존성
├── .env.example             # API 키 템플릿
├── docx_reader/
│   ├── omml_parser.py       # OMML XML → LaTeX 변환기
│   ├── extractor.py         # 문서 추출 (텍스트 + 수식)
│   └── splitter.py          # 문항 분리
├── gpt_solver/
│   ├── client.py            # OpenAI API 래퍼 (재시도 포함)
│   └── prompts.py           # 시스템/유저 프롬프트 템플릿
└── notebook_builder/
    └── builder.py           # nbformat 기반 .ipynb 생성
```

## 설치

```bash
pip install -r requirements.txt
```

## 설정

```bash
cp .env.example .env
```

`.env` 파일에 OpenAI API 키를 입력합니다:

```
OPENAI_API_KEY=sk-your-api-key-here
```

## 사용법

```bash
# 기본 실행 (입력 파일명 기반으로 .ipynb 자동 생성)
python main.py input.docx

# 출력 파일 지정
python main.py input.docx -o output.ipynb

# 모델 변경
python main.py input.docx -m gpt-4o

# API 호출 없이 수식 추출/문항 분리 결과만 확인
python main.py input.docx --dry-run
```

### 옵션

| 옵션 | 설명 | 기본값 |
|------|------|--------|
| `input` | 입력 Word 파일 경로 (.docx) | (필수) |
| `-o`, `--output` | 출력 노트북 파일 경로 (.ipynb) | `{입력파일명}.ipynb` |
| `-m`, `--model` | GPT 모델명 | `gpt-4o-mini` |
| `--dry-run` | API 호출 없이 추출 결과만 출력 | `false` |

## 의존성

- `python-docx>=1.1.0` — Word 문서 읽기
- `openai>=1.0.0` — ChatGPT API
- `nbformat>=5.7.0` — Jupyter Notebook 생성
- `python-dotenv>=1.0.0` — 환경변수 관리
