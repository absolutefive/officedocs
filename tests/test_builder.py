import io
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from officedocs.builder import build_deck
from officedocs.parser import parse_document

ROOT = Path(__file__).resolve().parent.parent
EXAMPLE = ROOT / "examples" / "sample.md"
TEMPLATES = ROOT / "templates"


def _build(template=None):
    deck = parse_document(EXAMPLE.read_text(encoding="utf-8"))
    prs = build_deck(deck, template, base_dir=EXAMPLE.parent)
    # 저장 → 재오픈 라운드트립으로 파일 유효성까지 확인
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)


def test_build_with_builtin_default():
    prs = _build()
    assert len(prs.slides) == 9  # 표지 1 + 본문 8


def test_titles_are_native_text():
    prs = _build()
    first = prs.slides[0]
    titles = [s for s in first.shapes if s.has_text_frame and "officedocs" in s.text_frame.text]
    assert titles, "표지 제목이 텍스트 개체로 존재해야 한다"


def test_table_is_native_object():
    prs = _build()
    tables = [
        shape
        for slide in prs.slides
        for shape in slide.shapes
        if shape.has_table
    ]
    assert tables, "표가 네이티브 표 개체여야 한다"
    assert tables[0].table.cell(0, 0).text == "기능"


def test_image_is_picture_shape():
    prs = _build()
    pictures = [
        shape
        for slide in prs.slides
        for shape in slide.shapes
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    ]
    assert pictures, "그림이 Picture 개체여야 한다"


def test_notes_round_trip():
    prs = _build()
    notes = [
        slide.notes_slide.notes_text_frame.text
        for slide in prs.slides
        if slide.has_notes_slide
    ]
    assert any("도입부" in n for n in notes)


def test_no_raster_only_output():
    """본문 텍스트가 이미지가 아니라 실제 텍스트 프레임에 존재하는지 확인."""
    prs = _build()
    all_text = "\n".join(
        shape.text_frame.text
        for slide in prs.slides
        for shape in slide.shapes
        if shape.has_text_frame
    )
    assert "내용은 마크다운 텍스트로 작성" in all_text


@pytest.mark.parametrize("name", ["default.pptx", "dark.pptx"])
def test_build_with_bundled_templates(name):
    template = TEMPLATES / name
    if not template.exists():
        pytest.skip("번들 템플릿이 아직 생성되지 않음 (scripts/make_templates.py)")
    prs = _build(template)
    assert len(prs.slides) == 9


TOMATO = TEMPLATES / "tomato" / "typeA-landscape.pptx"


@pytest.mark.parametrize(
    "name",
    ["typeA-landscape", "typeA-portrait", "typeB-landscape", "typeB-portrait"],
)
def test_build_with_prototype_templates(name):
    prs = _build(TEMPLATES / "tomato" / f"{name}.pptx")
    assert len(prs.slides) == 9


def test_prototype_markers_replaced_and_removed():
    prs = _build(TOMATO)
    all_text = "\n".join(
        shape.text_frame.text
        for slide in prs.slides
        for shape in slide.shapes
        if shape.has_text_frame
    )
    # 마커가 실제 내용으로 치환되어야 한다
    assert "officedocs 소개" in all_text
    assert "자리 배치" not in all_text
    # 안내문구 도형은 제거되어야 한다
    assert "표지타입" not in all_text
    assert "고객사 로고" not in all_text


def test_prototype_template_slides_stripped():
    """템플릿의 예시 슬라이드 4장이 결과물에 남지 않아야 한다."""
    deck = parse_document("# 한 장\n- 항목")
    prs = build_deck(deck, TOMATO, base_dir=EXAMPLE.parent)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    assert len(Presentation(buf).slides) == 1


def test_prototype_two_content_split():
    prs = _build(TOMATO)
    # sample.md의 '텍스트 vs 슬라이드' 슬라이드: 좌/우 텍스트 상자가 나란히 있어야 한다
    for slide in prs.slides:
        texts = [s for s in slide.shapes if s.has_text_frame and "git diff" in s.text_frame.text]
        if texts:
            left_box = texts[0]
            right = [
                s for s in slide.shapes
                if s.has_text_frame and "네이티브" in s.text_frame.text
            ]
            assert right and right[0].left > left_box.left
            return
    pytest.fail("two-content 슬라이드를 찾지 못함")


def test_prototype_table_and_image_native():
    prs = _build(TOMATO)
    assert any(sh.has_table for s in prs.slides for sh in s.shapes)
    from pptx.enum.shapes import MSO_SHAPE_TYPE as ST
    pics = [sh for s in prs.slides for sh in s.shapes if sh.shape_type == ST.PICTURE]
    # 본문 다이어그램 + 템플릿 로고들
    assert pics


AGENTOS_EXAMPLE = ROOT / "examples" / "agentos-sample.md"


def _build_agentos(colorway="teal"):
    deck = parse_document(AGENTOS_EXAMPLE.read_text(encoding="utf-8"))
    prs = build_deck(
        deck, TEMPLATES / "agentos" / f"{colorway}.pptx", base_dir=AGENTOS_EXAMPLE.parent
    )
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)


@pytest.mark.parametrize("colorway", ["teal", "blue", "red", "orange"])
def test_build_with_agentos_templates(colorway):
    prs = _build_agentos(colorway)
    assert len(prs.slides) == 7
    assert prs.slide_width == 12192000  # 13.33in 와이드


def test_code_window_rendered_as_shapes():
    prs = _build_agentos()
    from pptx.enum.shapes import MSO_SHAPE_TYPE as ST
    for slide in prs.slides:
        code_texts = [
            s for s in slide.shapes
            if s.has_text_frame and "Grill Me Skill" in s.text_frame.text
        ]
        if code_texts:
            autoshapes = [s for s in slide.shapes if s.shape_type == ST.AUTO_SHAPE]
            # 윈도우 + 헤더 + 신호등 점 3개 = 도형 5개 이상
            assert len(autoshapes) >= 5
            labels = [
                s for s in slide.shapes
                if s.has_text_frame and s.text_frame.text == "grill-me.md"
            ]
            assert labels, "파일명 라벨이 있어야 한다"
            return
    pytest.fail("코드 윈도우 슬라이드를 찾지 못함")


def test_eyebrow_rendered_above_title():
    prs = _build_agentos()
    for slide in prs.slides:
        eyebrows = [
            s for s in slide.shapes
            if s.has_text_frame and "6 TRAPS" in s.text_frame.text
        ]
        if eyebrows:
            titles = [
                s for s in slide.shapes
                if s.has_text_frame and "6가지 함정" in s.text_frame.text
            ]
            assert titles and eyebrows[0].top < titles[0].top
            return
    pytest.fail("아이브로우 슬라이드를 찾지 못함")


def test_missing_image_raises(tmp_path):
    doc = tmp_path / "doc.md"
    doc.write_text("# x\n![](no-such.png)", encoding="utf-8")
    deck = parse_document(doc.read_text(encoding="utf-8"))
    with pytest.raises(FileNotFoundError):
        build_deck(deck, base_dir=tmp_path)
