"""마크다운 기반 입력 템플릿 파서.

입력 형식 규칙 (자세한 예시는 examples/sample.md 참고):

- 문서 맨 위 ``---`` 로 감싼 머리말(frontmatter)에 ``key: value`` 메타데이터를
  적는다. ``title`` 이 있으면 표지(제목) 슬라이드가 자동 생성된다.
- 슬라이드는 단독 줄 ``---`` 로 구분한다.
- 슬라이드의 첫 ``#``/``##`` 제목이 슬라이드 제목이 된다.
- ``<!-- layout: NAME -->`` 으로 레이아웃을 지정한다.
  (title / section / content / two-content / title-only / blank)
- ``- `` 또는 ``* `` 글머리 기호, ``1. `` 번호 매기기, 2칸 들여쓰기로 하위 단계.
- ``| a | b |`` 파이프 표는 PPTX 네이티브 표가 된다.
- ``![대체텍스트](경로)`` 는 그림 개체가 된다.
- ``::: notes`` ~ ``:::`` 블록은 발표자 노트가 된다.
- ``::: left`` / ``::: right`` ~ ``:::`` 블록은 two-content 레이아웃의
  좌/우 영역에 배치된다.
- ``` ``` ``` 코드 펜스는 다크 코드 윈도우가 된다. 펜스 정보 문자열은
  파일명/언어 라벨로 표시된다 (예: ``` ```grill-me.md ``` ).
- ``<!-- eyebrow: TEXT -->`` 는 제목 위의 포인트 컬러 강조 라벨이 된다.
- 인라인 서식: ``**굵게**``, ``*기울임*``, `` `코드` ``.
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple

from officedocs.model import Block, CodeBlock, Deck, Image, Paragraph, Run, Slide, Table

_DIRECTIVE_RE = re.compile(r"^<!--\s*([\w-]+)\s*:\s*(.*?)\s*-->$")
_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
_BULLET_RE = re.compile(r"^(\s*)[-*]\s+(.*)$")
_NUMBERED_RE = re.compile(r"^(\s*)\d+[.)]\s+(.*)$")
_IMAGE_RE = re.compile(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$")
_TABLE_SEP_RE = re.compile(r"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$")
_FENCE_OPEN_RE = re.compile(r"^:::\s*(\w+)\s*$")
_INLINE_RE = re.compile(r"(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)")

VALID_LAYOUTS = {"title", "section", "content", "two-content", "title-only", "blank"}


def parse_inline(text: str) -> List[Run]:
    """``**굵게**``, ``*기울임*``, `` `코드` `` 를 Run 목록으로 분해한다."""
    runs: List[Run] = []
    for token in _INLINE_RE.split(text):
        if not token:
            continue
        if token.startswith("**") and token.endswith("**") and len(token) > 4:
            runs.append(Run(token[2:-2], bold=True))
        elif token.startswith("*") and token.endswith("*") and len(token) > 2:
            runs.append(Run(token[1:-1], italic=True))
        elif token.startswith("`") and token.endswith("`") and len(token) > 2:
            runs.append(Run(token[1:-1], code=True))
        else:
            runs.append(Run(token))
    return runs or [Run("")]


def _parse_frontmatter(lines: List[str]) -> Tuple[Dict[str, str], int]:
    """문서 머리의 ``--- ... ---`` 블록을 파싱한다. (메타, 본문 시작 줄) 반환."""
    idx = 0
    while idx < len(lines) and not lines[idx].strip():
        idx += 1
    if idx >= len(lines) or lines[idx].strip() != "---":
        return {}, 0
    meta: Dict[str, str] = {}
    for end in range(idx + 1, len(lines)):
        line = lines[end].strip()
        if line == "---":
            return meta, end + 1
        if line and ":" in line:
            key, _, value = line.partition(":")
            meta[key.strip().lower()] = value.strip()
    # 닫는 ---가 없으면 머리말이 아닌 것으로 간주
    return {}, 0


def _split_slides(lines: List[str]) -> List[List[str]]:
    chunks: List[List[str]] = []
    current: List[str] = []
    for line in lines:
        if line.strip() == "---":
            chunks.append(current)
            current = []
        else:
            current.append(line)
    chunks.append(current)
    return [c for c in chunks if any(l.strip() for l in c)]


def _parse_table(lines: List[str], start: int) -> Tuple[Table, int]:
    rows: List[List[List[Run]]] = []
    has_header = False
    i = start
    while i < len(lines) and lines[i].strip().startswith("|"):
        stripped = lines[i].strip()
        if _TABLE_SEP_RE.match(stripped):
            if len(rows) == 1:
                has_header = True
            i += 1
            continue
        # 셀 내용에 포함된 이스케이프 파이프(\|)는 분리 대상이 아님
        guarded = stripped.replace("\\|", "\x00")
        cells = [c.strip().replace("\x00", "|") for c in guarded.strip("|").split("|")]
        rows.append([parse_inline(c) for c in cells])
        i += 1
    return Table(rows=rows, has_header=has_header), i


def _parse_slide(lines: List[str]) -> Slide:
    slide = Slide()
    target: List[Block] = slide.blocks  # ::: left / ::: right 로 전환됨
    notes_lines: List[str] = []
    in_fence: Optional[str] = None
    i = 0
    while i < len(lines):
        raw = lines[i]
        stripped = raw.strip()

        if in_fence:
            if stripped == ":::":
                if in_fence == "notes":
                    slide.notes = "\n".join(notes_lines).strip()
                    notes_lines = []
                target = slide.blocks
                in_fence = None
            elif in_fence == "notes":
                notes_lines.append(raw.rstrip())
            else:
                # left/right 영역 내부는 일반 블록처럼 파싱하기 위해 통과시킨다
                consumed = _parse_block_line(lines, i, target)
                i = consumed
                continue
            i += 1
            continue

        fence = _FENCE_OPEN_RE.match(stripped)
        if fence:
            kind = fence.group(1).lower()
            if kind == "notes":
                in_fence = "notes"
            elif kind in ("left", "right"):
                in_fence = kind
                target = slide.left if kind == "left" else slide.right
                if slide.layout is None:
                    slide.layout = "two-content"
            i += 1
            continue

        directive = _DIRECTIVE_RE.match(stripped)
        if directive:
            key, value = directive.group(1).lower(), directive.group(2).lower()
            if key == "layout":
                if value not in VALID_LAYOUTS:
                    raise ValueError(
                        f"알 수 없는 레이아웃 '{value}'. 사용 가능: {sorted(VALID_LAYOUTS)}"
                    )
                slide.layout = value
            elif key == "eyebrow":
                slide.eyebrow = directive.group(2)  # 대소문자 유지
            i += 1
            continue

        heading = _HEADING_RE.match(stripped)
        if heading and slide.title is None:
            slide.title = heading.group(2).strip()
            i += 1
            continue

        i = _parse_block_line(lines, i, target)

    return slide


def _parse_block_line(lines: List[str], i: int, target: List[Block]) -> int:
    """lines[i]부터 블록 하나를 파싱해 target에 추가하고 다음 인덱스를 반환."""
    raw = lines[i]
    stripped = raw.strip()

    if not stripped:
        return i + 1

    if stripped.startswith("```"):
        label = stripped[3:].strip()
        code_lines: List[str] = []
        j = i + 1
        while j < len(lines) and lines[j].strip() != "```":
            code_lines.append(lines[j].rstrip())
            j += 1
        target.append(CodeBlock(lines=code_lines, label=label))
        return j + 1  # 닫는 펜스 건너뜀 (없으면 슬라이드 끝까지)

    image = _IMAGE_RE.match(stripped)
    if image:
        target.append(Image(path=image.group(2).strip(), alt=image.group(1).strip()))
        return i + 1

    if stripped.startswith("|"):
        table, next_i = _parse_table(lines, i)
        if table.rows:
            target.append(table)
        return next_i

    bullet = _BULLET_RE.match(raw)
    if bullet:
        level = min(len(bullet.group(1).expandtabs(2)) // 2, 4)
        target.append(Paragraph(parse_inline(bullet.group(2)), level=level, bullet=True))
        return i + 1

    numbered = _NUMBERED_RE.match(raw)
    if numbered:
        level = min(len(numbered.group(1).expandtabs(2)) // 2, 4)
        target.append(
            Paragraph(parse_inline(numbered.group(2)), level=level, numbered=True)
        )
        return i + 1

    heading = _HEADING_RE.match(stripped)
    text = heading.group(2).strip() if heading else stripped
    target.append(Paragraph(parse_inline(text)))
    return i + 1


def parse_document(text: str) -> Deck:
    """입력 텍스트 전체를 Deck 모델로 파싱한다."""
    lines = text.splitlines()
    meta, body_start = _parse_frontmatter(lines)
    deck = Deck(meta=meta)

    if meta.get("title"):
        cover = Slide(layout="title", title=meta["title"])
        subtitle_parts = [meta.get(k, "") for k in ("subtitle", "author", "date")]
        for part in subtitle_parts:
            if part:
                cover.blocks.append(Paragraph(parse_inline(part)))
        deck.slides.append(cover)

    for chunk in _split_slides(lines[body_start:]):
        deck.slides.append(_parse_slide(chunk))

    return deck
