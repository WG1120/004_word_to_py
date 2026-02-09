# 실행 가이드

## 1. 의존성 설치

```bash
cd C:\Users\wnsdn\v_code\004_word_to_py
pip install -r requirements.txt
```

## 2. OpenAI API 키 발급

1. https://platform.openai.com 접속 후 로그인
2. 좌측 메뉴에서 **API keys** 클릭 (또는 https://platform.openai.com/api-keys 직접 접속)
3. **"Create new secret key"** 클릭
4. 이름을 입력하고 생성하면 `sk-...` 형태의 키가 표시됨
5. **이 키를 반드시 복사해두세요** (창을 닫으면 다시 볼 수 없음)

> 참고: API 사용에는 크레딧이 필요합니다. https://platform.openai.com/settings/organization/billing 에서 결제 수단을 등록하고 크레딧을 충전하세요. gpt-4o-mini는 매우 저렴합니다.

## 3. API 키 설정

프로젝트 폴더에 `.env` 파일을 생성합니다:

```bash
copy .env.example .env
```

그리고 `.env` 파일을 열어 발급받은 키를 입력합니다:

```
OPENAI_API_KEY=sk-proj-xxxxxxxxxxxxxxxxxxxxxxxx
```

## 4. 실행

```bash
# 먼저 dry-run으로 수식 추출/문항 분리가 잘 되는지 확인 (API 호출 안 함, 무료)
python main.py sample.docx --dry-run

# 확인 후 전체 실행 (API 호출하여 풀이 생성 + 노트북 출력)
python main.py sample.docx

# 출력 파일명 지정
python main.py sample.docx -o 풀이결과.ipynb

# 다른 모델 사용
python main.py sample.docx -m gpt-4o
```

## 5. 결과 확인

생성된 `.ipynb` 파일을 Jupyter Notebook에서 엽니다:

```bash
jupyter notebook sample.ipynb
```

또는 VS Code에서 `.ipynb` 파일을 직접 열어도 됩니다.

## 실행 흐름 요약

```
sample.docx
  → [1단계] 텍스트 + 수식(LaTeX) 추출
  → [2단계] 문항 분리
  → [3단계] GPT API로 각 문항 풀이 코드 생성
  → [4단계] Jupyter Notebook(.ipynb) 파일 저장
```
