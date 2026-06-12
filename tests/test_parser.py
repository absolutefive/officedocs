from officedocs.model import Image, Paragraph, Table
from officedocs.parser import parse_document, parse_inline

SAMPLE = """\
---
title: 제목입니다
subtitle: 부제
author: 홍길동
---

<!-- layout: section -->
# 섹션 하나

---

# 내용 슬라이드
설명 문단입니다.

- 항목 1
  - 하위 항목
- **굵은** 항목

::: notes
노트 첫 줄
노트 둘째 줄
:::

---

# 표 슬라이드
| 이름 | 값 |
| --- | --- |
| a | 1 |
| b | 2 |

---

<!-- layout: two-content -->
# 좌우 비교

::: left
- 왼쪽 항목
:::

::: right
- 오른쪽 항목
:::

---

# 그림
![구조도](assets/diagram.png)
"""


def test_frontmatter_creates_cover_slide():
    deck = parse_document(SAMPLE)
    assert deck.meta["title"] == "제목입니다"
    cover = deck.slides[0]
    assert cover.layout == "title"
    assert cover.title == "제목입니다"
    subtitle_texts = [b.text for b in cover.blocks]
    assert subtitle_texts == ["부제", "홍길동"]


def test_section_directive():
    deck = parse_document(SAMPLE)
    section = deck.slides[1]
    assert section.layout == "section"
    assert section.title == "섹션 하나"


def test_bullets_nesting_and_inline():
    deck = parse_document(SAMPLE)
    slide = deck.slides[2]
    assert slide.title == "내용 슬라이드"
    bullets = [b for b in slide.blocks if isinstance(b, Paragraph) and b.bullet]
    assert [b.level for b in bullets] == [0, 1, 0]
    bold_runs = [r for r in bullets[2].runs if r.bold]
    assert bold_runs and bold_runs[0].text == "굵은"


def test_notes():
    deck = parse_document(SAMPLE)
    assert deck.slides[2].notes == "노트 첫 줄\n노트 둘째 줄"


def test_table():
    deck = parse_document(SAMPLE)
    tables = [b for b in deck.slides[3].blocks if isinstance(b, Table)]
    assert len(tables) == 1
    table = tables[0]
    assert table.has_header
    assert len(table.rows) == 3
    assert table.rows[1][0][0].text == "a"


def test_two_content_columns():
    deck = parse_document(SAMPLE)
    slide = deck.slides[4]
    assert slide.layout == "two-content"
    assert slide.left[0].text == "왼쪽 항목"
    assert slide.right[0].text == "오른쪽 항목"


def test_image():
    deck = parse_document(SAMPLE)
    images = [b for b in deck.slides[5].blocks if isinstance(b, Image)]
    assert images[0].path == "assets/diagram.png"


def test_parse_inline_mixed():
    runs = parse_inline("일반 **굵게** 그리고 *기울임* 과 `code` 끝")
    kinds = [(r.bold, r.italic, r.code) for r in runs]
    assert (True, False, False) in kinds
    assert (False, True, False) in kinds
    assert (False, False, True) in kinds
    assert "".join(r.text for r in runs) == "일반 굵게 그리고 기울임 과 code 끝"


def test_invalid_layout_raises():
    import pytest

    with pytest.raises(ValueError):
        parse_document("<!-- layout: nonsense -->\n# x")


def test_no_frontmatter():
    deck = parse_document("# 단독 슬라이드\n- 항목")
    assert deck.meta == {}
    assert len(deck.slides) == 1
    assert deck.slides[0].title == "단독 슬라이드"


def test_escaped_pipe_in_table_cell():
    deck = parse_document("# t\n| 문법 |\n| --- |\n| `\\| a \\|` |")
    table = [b for b in deck.slides[0].blocks if isinstance(b, Table)][0]
    assert table.rows[1][0][0].text == "| a |"
