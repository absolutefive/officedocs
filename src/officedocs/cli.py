"""officedocs 명령줄 인터페이스.

사용 예::

    officedocs build deck.md -o out.pptx                 # 기본 템플릿
    officedocs build deck.md -t templates/dark.pptx      # 템플릿 교체 후 재생성
    officedocs layouts templates/dark.pptx               # 템플릿 레이아웃 확인
    officedocs init my-deck.md                           # 입력 문서 스캐폴드 생성
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pptx import Presentation

from officedocs import __version__
from officedocs.builder import build_deck
from officedocs.layouts import LayoutResolver
from officedocs.parser import parse_document

_SCAFFOLD = """\
---
title: 발표 제목
subtitle: 부제목
author: 작성자
date: 2026-01-01
---

<!-- layout: section -->
# 첫 번째 섹션

---

# 내용 슬라이드
- 첫 번째 항목
  - 하위 항목
- **굵은** 항목과 *기울임* 항목

::: notes
발표자 노트는 여기에 적습니다.
:::

---

# 표 예시
| 항목 | 값 |
| --- | --- |
| 속도 | 빠름 |
| 비용 | 낮음 |
"""


def _cmd_build(args: argparse.Namespace) -> int:
    input_path = Path(args.input)
    text = input_path.read_text(encoding="utf-8")
    deck = parse_document(text)

    template = args.template or deck.meta.get("template")
    template_path = None
    if template:
        template_path = Path(template)
        if not template_path.is_absolute() and not template_path.exists():
            candidate = input_path.parent / template_path
            if candidate.exists():
                template_path = candidate
        if not template_path.exists():
            print(f"오류: 템플릿 파일이 없습니다: {template}", file=sys.stderr)
            return 1

    output = Path(args.output) if args.output else input_path.with_suffix(".pptx")
    prs = build_deck(deck, template_path, base_dir=input_path.parent)
    prs.save(str(output))
    print(f"생성 완료: {output} (슬라이드 {len(prs.slides)}장, "
          f"템플릿: {template_path or '기본 내장'})")
    return 0


def _cmd_layouts(args: argparse.Namespace) -> int:
    template_path = Path(args.template)
    prs = Presentation(str(template_path))
    resolver = LayoutResolver(prs, template_path)
    print(resolver.describe())
    return 0


def _cmd_init(args: argparse.Namespace) -> int:
    output = Path(args.output)
    if output.exists():
        print(f"오류: 이미 파일이 존재합니다: {output}", file=sys.stderr)
        return 1
    output.write_text(_SCAFFOLD, encoding="utf-8")
    print(f"입력 문서 생성: {output}")
    return 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="officedocs", description="텍스트 문서로부터 편집 가능한 PPTX 생성"
    )
    parser.add_argument("--version", action="version", version=f"officedocs {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    p_build = sub.add_parser("build", help="입력 문서를 PPTX로 빌드")
    p_build.add_argument("input", help="입력 텍스트 문서(.md)")
    p_build.add_argument("-t", "--template", help="PPTX 템플릿 파일 경로")
    p_build.add_argument("-o", "--output", help="출력 PPTX 경로 (기본: 입력명.pptx)")
    p_build.set_defaults(func=_cmd_build)

    p_layouts = sub.add_parser("layouts", help="템플릿의 레이아웃/매핑 확인")
    p_layouts.add_argument("template", help="PPTX 템플릿 파일 경로")
    p_layouts.set_defaults(func=_cmd_layouts)

    p_init = sub.add_parser("init", help="입력 문서 스캐폴드 생성")
    p_init.add_argument("output", help="생성할 입력 문서 경로")
    p_init.set_defaults(func=_cmd_init)

    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
