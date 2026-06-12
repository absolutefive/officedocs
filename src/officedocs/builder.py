"""Deck 모델 + PPTX 템플릿 → 최종 PPTX 빌더.

모든 내용은 자리표시자/텍스트 상자/표/그림 등 네이티브 PowerPoint 개체로
생성되므로, 결과물을 PowerPoint에서 직접 수정할 수 있다.
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Emu, Pt

from officedocs.layouts import LayoutResolver
from officedocs.model import Block, Deck, Image, Paragraph, Run, Slide, Table

_TITLE_TYPES = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
_BODY_TYPES = (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)
_CODE_FONT = "Consolas"


def _find_placeholder(slide, types, skip=0):
    """slide에서 주어진 형식의 자리표시자를 idx 순서로 찾는다."""
    found = sorted(
        (ph for ph in slide.placeholders if ph.placeholder_format.type in types),
        key=lambda ph: ph.placeholder_format.idx,
    )
    return found[skip] if len(found) > skip else None


def _apply_runs(paragraph, runs: List[Run]):
    for r in runs:
        run = paragraph.add_run()
        run.text = r.text
        run.font.bold = r.bold or None
        run.font.italic = r.italic or None
        if r.code:
            run.font.name = _CODE_FONT


def _fill_text_frame(text_frame, paragraphs: List[Paragraph]):
    """자리표시자의 텍스트 프레임에 문단들을 채운다."""
    text_frame.clear()
    for i, para in enumerate(paragraphs):
        p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        p.level = para.level
        _apply_runs(p, para.runs)


class DeckBuilder:
    def __init__(self, deck: Deck, template_path: Optional[Path] = None,
                 base_dir: Optional[Path] = None):
        self._deck = deck
        self._template_path = template_path
        self._base_dir = base_dir or Path.cwd()
        self._prs = (
            Presentation(str(template_path)) if template_path else Presentation()
        )
        self._strip_existing_slides()
        self._resolver = LayoutResolver(self._prs, template_path)

    def _strip_existing_slides(self):
        """템플릿 파일에 남아있는 예시 슬라이드를 제거한다(레이아웃은 유지)."""
        xml_slides = self._prs.slides._sldIdLst
        for slide_id in list(xml_slides):
            xml_slides.remove(slide_id)

    def build(self) -> Presentation:
        for spec in self._deck.slides:
            self._add_slide(spec)
        return self._prs

    # --- 슬라이드 생성 ---------------------------------------------------

    def _add_slide(self, spec: Slide):
        layout_name = spec.layout or self._auto_layout(spec)
        layout = self._resolver.get(layout_name)
        slide = self._prs.slides.add_slide(layout)

        if spec.title is not None:
            title_ph = _find_placeholder(slide, _TITLE_TYPES)
            if title_ph is not None:
                title_ph.text_frame.text = spec.title

        if layout_name == "title":
            self._fill_title_slide(slide, spec)
        elif spec.left or spec.right:
            self._fill_body(slide, spec.left, body_skip=0)
            self._fill_body(slide, spec.right, body_skip=1)
            if spec.blocks:
                self._place_floating_blocks(slide, spec.blocks)
        else:
            self._fill_body(slide, spec.blocks, body_skip=0)

        if spec.notes:
            slide.notes_slide.notes_text_frame.text = spec.notes

    @staticmethod
    def _auto_layout(spec: Slide) -> str:
        if spec.left or spec.right:
            return "two-content"
        if spec.title and not spec.blocks:
            return "title-only"
        if not spec.title and not spec.blocks:
            return "blank"
        return "content"

    def _fill_title_slide(self, slide, spec: Slide):
        subtitle = _find_placeholder(slide, (PP_PLACEHOLDER.SUBTITLE,)) or _find_placeholder(
            slide, _BODY_TYPES
        )
        paragraphs = [b for b in spec.blocks if isinstance(b, Paragraph)]
        if subtitle is not None and paragraphs:
            _fill_text_frame(subtitle.text_frame, paragraphs)

    def _fill_body(self, slide, blocks: List[Block], body_skip: int):
        if not blocks:
            return
        body = _find_placeholder(slide, _BODY_TYPES, skip=body_skip)
        text_blocks = [b for b in blocks if isinstance(b, Paragraph)]
        rich_blocks = [b for b in blocks if not isinstance(b, Paragraph)]

        if body is not None and text_blocks:
            _fill_text_frame(body.text_frame, text_blocks)

        if rich_blocks:
            area = self._block_area(slide, body, below_text=bool(text_blocks))
            self._place_rich_blocks(slide, rich_blocks, area)
        elif body is None and text_blocks:
            # 본문 자리표시자가 없는 레이아웃(title-only/blank)이면 텍스트 상자 생성
            area = self._default_area(slide)
            box = slide.shapes.add_textbox(*area)
            box.text_frame.word_wrap = True
            _fill_text_frame(box.text_frame, text_blocks)

    # --- 표/그림 배치 -----------------------------------------------------

    def _default_area(self, slide):
        sw, sh = self._prs.slide_width, self._prs.slide_height
        margin = Emu(int(sw * 0.06))
        top = Emu(int(sh * 0.25))
        return margin, top, Emu(sw - 2 * margin), Emu(sh - top - margin)

    def _block_area(self, slide, body, below_text: bool):
        if body is None:
            return self._default_area(slide)
        left, top = body.left, body.top
        width, height = body.width, body.height
        if below_text:
            # 텍스트가 위쪽 절반을 차지한다고 보고 아래 절반에 배치
            top = Emu(int(top + height * 0.55))
            height = Emu(int(height * 0.45))
        return left, top, width, height

    def _place_floating_blocks(self, slide, blocks: List[Block]):
        self._place_rich_blocks(
            slide,
            [b for b in blocks if not isinstance(b, Paragraph)],
            self._default_area(slide),
        )

    def _place_rich_blocks(self, slide, blocks: List[Block], area):
        left, top, width, height = area
        cursor = top
        per_block = Emu(int(height / max(len(blocks), 1)))
        for block in blocks:
            if isinstance(block, Table):
                self._add_table(slide, block, left, cursor, width)
            elif isinstance(block, Image):
                self._add_image(slide, block, left, cursor, width, per_block)
            cursor = Emu(int(cursor + per_block))

    def _add_table(self, slide, table: Table, left, top, width):
        n_rows = len(table.rows)
        n_cols = max(len(r) for r in table.rows)
        shape = slide.shapes.add_table(
            n_rows, n_cols, left, top, width, Emu(Pt(24) * n_rows)
        )
        for ri, row in enumerate(table.rows):
            for ci in range(n_cols):
                cell = shape.table.cell(ri, ci)
                runs = row[ci] if ci < len(row) else [Run("")]
                tf = cell.text_frame
                tf.clear()
                _apply_runs(tf.paragraphs[0], runs)
                if ri == 0 and table.has_header:
                    for run in tf.paragraphs[0].runs:
                        run.font.bold = True

    def _add_image(self, slide, image: Image, left, top, width, max_height):
        path = Path(image.path)
        if not path.is_absolute():
            path = self._base_dir / path
        if not path.exists():
            raise FileNotFoundError(f"이미지 파일을 찾을 수 없습니다: {image.path}")
        pic = slide.shapes.add_picture(str(path), left, top, width=width)
        if pic.height > max_height:
            # 영역을 넘으면 높이에 맞춰 비율 유지 축소
            scale = max_height / pic.height
            pic.width = Emu(int(pic.width * scale))
            pic.height = Emu(int(pic.height * scale))


def build_deck(deck: Deck, template_path: Optional[Path] = None,
               base_dir: Optional[Path] = None) -> Presentation:
    """Deck 모델을 PPTX Presentation으로 빌드한다."""
    return DeckBuilder(deck, template_path, base_dir).build()
