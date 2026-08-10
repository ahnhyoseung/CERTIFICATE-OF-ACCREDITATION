"""
Claude와 실시간으로 '대화하듯' 주고받는 간단한 콘솔 챗봇.

- pip install anthropic
- 환경변수 ANTHROPIC_API_KEY 설정 필요

사용법:
    python chat_with_claude.py

종료: exit / quit 입력, 또는 Ctrl+C
"""

import os
import sys

import anthropic

MODEL = "claude-sonnet-4-6"

# 필요하면 여기에 시스템 프롬프트(역할 지정)를 넣을 수 있습니다.
SYSTEM_PROMPT = "당신은 친절하고 정확한 한국어 어시스턴트입니다."


def main() -> None:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("[에러] ANTHROPIC_API_KEY 환경변수가 설정되어 있지 않습니다.", file=sys.stderr)
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    # 대화 기록을 여기에 계속 쌓아나갑니다. 이게 '기억'의 전부입니다.
    history = []

    print("Claude와 대화를 시작합니다. 종료하려면 'exit' 또는 'quit' 입력.\n")

    while True:
        try:
            user_input = input("나: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n대화를 종료합니다.")
            break

        if user_input.lower() in ("exit", "quit"):
            print("대화를 종료합니다.")
            break
        if not user_input:
            continue

        # 사용자 메시지를 기록에 추가
        history.append({"role": "user", "content": user_input})

        try:
            response = client.messages.create(
                model=MODEL,
                max_tokens=1024,
                system=SYSTEM_PROMPT,
                messages=history,  # 지금까지의 대화 전체를 매번 다시 보냄
            )
        except Exception as e:
            print(f"[에러] API 호출 실패: {e}", file=sys.stderr)
            # 실패한 사용자 메시지는 기록에서 제거해서 다음 턴에 꼬이지 않게 함
            history.pop()
            continue

        answer = "".join(
            block.text for block in response.content if block.type == "text"
        ).strip()

        print(f"Claude: {answer}\n")

        # Claude의 답변도 기록에 추가해야 다음 턴에서 문맥을 이어갑니다.
        history.append({"role": "assistant", "content": answer})


if __name__ == "__main__":
    main()