"""Word to Jupyter Notebook Solver - CLI 진입점.

사용법:
    python main.py input.docx [-o output.ipynb] [-m gpt-4o-mini] [--dry-run]
"""

import argparse
import sys
from pathlib import Path

from docx_reader import extract_document, split_questions
from gpt_solver import solve_question
from notebook_builder import build_notebook
from notebook_builder.builder import SolvedQuestion
import config


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Word 문서의 데이터분석 문항을 GPT로 풀어 Jupyter Notebook으로 출력합니다.",
    )
    parser.add_argument(
        "input",
        help="입력 Word 파일 경로 (.docx)",
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="출력 노트북 파일 경로 (.ipynb). 미지정 시 입력 파일명 기반 자동 생성",
    )
    parser.add_argument(
        "-m", "--model",
        default=config.DEFAULT_MODEL,
        help=f"사용할 GPT 모델 (기본: {config.DEFAULT_MODEL})",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="API 호출 없이 문서 추출/문항 분리 결과만 출력",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"오류: 파일을 찾을 수 없습니다: {args.input}", file=sys.stderr)
        sys.exit(1)

    # 출력 경로 설정
    if args.output:
        output_path = args.output
    else:
        output_path = str(input_path.with_suffix(".ipynb"))

    total_steps = 2 if args.dry_run else 4

    # ── 1단계: 문서 추출 ──
    print(f"[1/{total_steps}] 문서에서 텍스트와 수식을 추출합니다...")
    try:
        full_text = extract_document(str(input_path))
    except Exception as e:
        print(f"오류: 문서 추출 실패 - {e}", file=sys.stderr)
        sys.exit(1)

    print(f"  추출 완료: {len(full_text)}자")

    # ── 2단계: 문항 분리 ──
    print(f"[2/{total_steps}] 문항을 분리합니다...")
    questions = split_questions(full_text)
    print(f"  {len(questions)}개 문항 감지")

    for q in questions:
        preview = q.body[:80].replace("\n", " ")
        print(f"  - 문제 {q.number}: {preview}...")

    # dry-run 모드
    if args.dry_run:
        print("\n=== Dry-run 모드: 추출 결과 ===\n")
        print(full_text)
        print(f"\n=== {len(questions)}개 문항 분리 완료 ===")
        return

    # ── 3단계: GPT 풀이 생성 ──
    print(f"[3/{total_steps}] GPT ({args.model})로 풀이를 생성합니다...")
    solved = []
    for i, q in enumerate(questions, 1):
        print(f"  [{i}/{len(questions)}] 문제 {q.number} 풀이 중...")
        try:
            code = solve_question(q.body, model=args.model)
            solved.append(SolvedQuestion(
                number=q.number,
                title=q.title,
                body=q.body,
                code=code,
            ))
            print(f"  [{i}/{len(questions)}] 문제 {q.number} 완료 ({len(code)}자)")
        except Exception as e:
            print(f"  경고: 문제 {q.number} 풀이 실패 - {e}", file=sys.stderr)
            solved.append(SolvedQuestion(
                number=q.number,
                title=q.title,
                body=q.body,
                code=f"# 풀이 생성 실패: {e}",
            ))

    # ── 4단계: 노트북 생성 ──
    print(f"[4/{total_steps}] Jupyter Notebook을 생성합니다...")
    title = f"{input_path.stem} - 풀이"
    build_notebook(solved, title=title, output_path=output_path)
    print(f"  저장 완료: {output_path}")
    print(f"\n완료! {len(solved)}개 문항의 풀이가 생성되었습니다.")


if __name__ == "__main__":
    main()
