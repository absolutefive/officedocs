"""AgentOS 스타일 템플릿 4종(templates/agentos/*.pptx) 생성 스크립트.

업로드된 ``gen_deck.js`` (pptxgenjs 기반 AgentOS 스타일 생성기)의 디자인
토큰을 officedocs 템플릿으로 이식했다. 화이트 배경 / 포인트 컬러 /
Pretendard 초고굵기 타이틀 + 검정 대시 / 자간 넓은 아이브로우 스타일.

원본 컬러웨이 4종(teal/blue/red/orange)이 그대로 재현되며, 컬러는 테마
색상으로 정의되어 있어 PALETTES만 수정하면 새 컬러웨이를 추가할 수 있다.

실행::

    python scripts/make_agentos_templates.py
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "templates" / "agentos"

# ── gen_deck.js 공통 토큰 ──────────────────────────────────────────
INK = "1B1C1E"     # 타이틀 (거의 검정) → 테마 dk1
BODY = "3E434B"    # 본문 → 테마 dk2
MUTED = "8A8F98"   # 보조 회색
CARD_GRAY = "F4F4F5"  # 중립 카드 배경 → 테마 lt2
F_DISP = "Pretendard ExtraBold"
F_MAIN = "Pretendard"

# ── gen_deck.js 컬러웨이 4종 ──────────────────────────────────────
PALETTES = [
    {"id": "teal", "ACCENT": "0F8A73", "TINT": "EBF6F2", "CODE_HL": "9FD8A8"},
    {"id": "blue", "ACCENT": "2563EB", "TINT": "EBF1FE", "CODE_HL": "9DBDF9"},
    {"id": "red", "ACCENT": "DC2626", "TINT": "FCEDED", "CODE_HL": "F2A6A6"},
    {"id": "orange", "ACCENT": "EA580C", "TINT": "FDF0E6", "CODE_HL": "F6BE8F"},
]

_A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
_META_PH_TYPES = (
    PP_PLACEHOLDER.DATE,
    PP_PLACEHOLDER.FOOTER,
    PP_PLACEHOLDER.SLIDE_NUMBER,
)


def _lst_style(levels) -> str:
    """레이아웃 자리표시자용 <a:lstStyle> XML을 만든다."""
    out = [f"<a:lstStyle {_A_NS}>"]
    for i, lv in enumerate(levels, start=1):
        attrs = ""
        if lv.get("algn"):
            attrs += f' algn="{lv["algn"]}"'
        if lv.get("bullet"):
            mar = 274320 * i
            attrs += f' marL="{mar}" indent="-228600"'
        out.append(f"<a:lvl{i}pPr{attrs}>")
        if lv.get("space_before"):
            out.append(f'<a:spcBef><a:spcPts val="{lv["space_before"]}"/></a:spcBef>')
        if lv.get("bullet"):
            # 글머리 기호는 포인트 컬러(테마 강조1)
            out.append(
                '<a:buClr><a:schemeClr val="accent1"/></a:buClr>'
                '<a:buFont typeface="Arial" pitchFamily="34" charset="0"/>'
                '<a:buChar char="&#8226;"/>'
            )
        else:
            out.append("<a:buNone/>")
        bold = ' b="1"' if lv.get("bold") else ""
        out.append(f'<a:defRPr sz="{lv["sz"]}"{bold}>')
        color = lv.get("color")
        if color:
            if color.startswith("scheme:"):
                out.append(f'<a:solidFill><a:schemeClr val="{color[7:]}"/></a:solidFill>')
            else:
                out.append(f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>')
        if lv.get("font"):
            out.append(f'<a:latin typeface="{lv["font"]}"/><a:ea typeface="{lv["font"]}"/>')
        out.append(f"</a:defRPr></a:lvl{i}pPr>")
    out.append("</a:lstStyle>")
    return "".join(out)


def _apply_style(ph, x, y, w, h, levels):
    ph.left, ph.top = Inches(x), Inches(y)
    ph.width, ph.height = Inches(w), Inches(h)
    txBody = ph.text_frame._txBody
    new = parse_xml(_lst_style(levels))
    old = txBody.find(qn("a:lstStyle"))
    if old is not None:
        txBody.replace(old, new)
    else:
        txBody.insert(1, new)  # bodyPr 다음


def _ph(layout, types, skip=0):
    found = [p for p in layout.placeholders if p.placeholder_format.type in types]
    found.sort(key=lambda p: p.placeholder_format.idx)
    return found[skip] if len(found) > skip else None


def _remove_meta_placeholders(layout):
    for p in list(layout.placeholders):
        if p.placeholder_format.type in _META_PH_TYPES:
            p._element.getparent().remove(p._element)


def _add_dash(layout, scratch):
    """타이틀 아래 짧은 검정 대시 (테마 텍스트1 색 → 컬러웨이와 무관하게 잉크색).

    레이아웃에는 add_shape가 없으므로 임시 슬라이드에 만들어 XML을 이식한다.
    """
    dash = scratch.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.56), Inches(1.68), Inches(0.55), Inches(0.045)
    )
    dash.fill.solid()
    dash.fill.fore_color.theme_color = MSO_THEME_COLOR.TEXT_1
    dash.line.fill.background()
    dash.shadow.inherit = False
    layout.shapes._spTree.append(dash._element)


_TITLE_TYPES = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
_BODY_TYPES = (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)

_BODY_LEVELS = [
    {"sz": 1500, "bullet": True, "color": BODY, "font": F_MAIN, "space_before": 900},
    {"sz": 1300, "bullet": True, "color": BODY, "font": F_MAIN, "space_before": 500},
    {"sz": 1200, "bullet": True, "color": MUTED, "font": F_MAIN, "space_before": 400},
]


def build_base() -> bytes:
    prs = Presentation()
    prs.slide_width = Emu(12192000)   # 13.333 x 7.5in (LAYOUT_WIDE)
    prs.slide_height = Emu(6858000)
    layouts = prs.slide_masters[0].slide_layouts
    scratch = prs.slides.add_slide(layouts[6])  # 도형 제작용 임시 슬라이드

    # [0] Title Slide — 센터 커버
    lay = layouts[0]
    _remove_meta_placeholders(lay)
    _apply_style(lay.placeholders[0], 0.5, 2.5, 12.33, 1.0,
                 [{"sz": 4400, "bold": True, "algn": "ctr", "color": "scheme:tx1", "font": F_DISP}])
    _apply_style(_ph(lay, (PP_PLACEHOLDER.SUBTITLE,)), 0.5, 3.75, 12.33, 1.1,
                 [{"sz": 1500, "algn": "ctr", "color": MUTED, "font": F_MAIN}])

    # [1] Title and Content — 좌상단 타이틀 + 대시 + 본문
    lay = layouts[1]
    _remove_meta_placeholders(lay)
    _apply_style(_ph(lay, _TITLE_TYPES), 0.53, 0.62, 12.25, 0.9,
                 [{"sz": 3400, "bold": True, "color": "scheme:tx1", "font": F_DISP}])
    _apply_style(_ph(lay, _BODY_TYPES), 0.55, 1.95, 12.23, 4.95, _BODY_LEVELS)
    _add_dash(lay, scratch)

    # [2] Section Header — 센터 인용형
    lay = layouts[2]
    _remove_meta_placeholders(lay)
    _apply_style(_ph(lay, _TITLE_TYPES), 0.5, 2.78, 12.33, 1.0,
                 [{"sz": 4200, "bold": True, "algn": "ctr", "color": "scheme:tx1", "font": F_DISP}])
    _apply_style(_ph(lay, _BODY_TYPES), 0.5, 3.95, 12.33, 0.9,
                 [{"sz": 1500, "algn": "ctr", "color": MUTED, "font": F_MAIN}])
    quote = scratch.shapes.add_textbox(Inches(0), Inches(2.3), Inches(13.33), Inches(0.45))
    quote.text_frame.text = "“"
    p = quote.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    for run in p.runs:
        run.font.name = F_DISP
        run.font.bold = True
        run.font.size = Pt(22)
        run.font.color.theme_color = MSO_THEME_COLOR.TEXT_1
    lay.shapes._spTree.append(quote._element)

    # [3] Two Content — 타이틀 + 대시 + 좌/우 본문
    lay = layouts[3]
    _remove_meta_placeholders(lay)
    _apply_style(_ph(lay, _TITLE_TYPES), 0.53, 0.62, 12.25, 0.9,
                 [{"sz": 3400, "bold": True, "color": "scheme:tx1", "font": F_DISP}])
    _apply_style(_ph(lay, _BODY_TYPES, 0), 0.55, 1.95, 6.0, 4.95, _BODY_LEVELS)
    _apply_style(_ph(lay, _BODY_TYPES, 1), 6.78, 1.95, 6.0, 4.95, _BODY_LEVELS)
    _add_dash(lay, scratch)

    # [5] Title Only
    lay = layouts[5]
    _remove_meta_placeholders(lay)
    _apply_style(_ph(lay, _TITLE_TYPES), 0.53, 0.62, 12.25, 0.9,
                 [{"sz": 3400, "bold": True, "color": "scheme:tx1", "font": F_DISP}])
    _add_dash(lay, scratch)

    # [6] Blank
    _remove_meta_placeholders(layouts[6])

    # 임시 슬라이드 제거
    sldIdLst = prs.slides._sldIdLst
    sldId = list(sldIdLst)[0]
    prs.part.drop_rel(sldId.get(qn("r:id")))
    sldIdLst.remove(sldId)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def apply_palette(base: bytes, palette: dict) -> bytes:
    """테마 XML의 색상/글꼴을 컬러웨이에 맞게 치환한다."""
    colors = {
        "dk1": INK,
        "lt1": "FFFFFF",
        "dk2": BODY,
        "lt2": CARD_GRAY,
        "accent1": palette["ACCENT"],
        "accent2": palette["TINT"],
        "accent3": palette["CODE_HL"],
        "accent4": MUTED,
        "hlink": palette["ACCENT"],
        "folHlink": MUTED,
    }
    with zipfile.ZipFile(io.BytesIO(base)) as zf:
        items = {name: zf.read(name) for name in zf.namelist()}
    for name in [n for n in items if re.match(r"ppt/theme/theme\d+\.xml", n)]:
        xml = items[name].decode("utf-8")
        for slot, color in colors.items():
            xml = re.sub(
                rf"(<a:{slot}>).*?(</a:{slot}>)",
                rf'\1<a:srgbClr val="{color}"/>\2',
                xml,
                flags=re.DOTALL,
            )
        xml = xml.replace('typeface="Calibri Light"', f'typeface="{F_MAIN}"')
        xml = xml.replace('typeface="Calibri"', f'typeface="{F_MAIN}"')
        items[name] = xml.encode("utf-8")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in items.items():
            zf.writestr(name, data)
    return out.getvalue()


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    base = build_base()
    for palette in PALETTES:
        path = OUT_DIR / f"{palette['id']}.pptx"
        path.write_bytes(apply_palette(base, palette))
        Presentation(str(path))  # 손상 여부 검증
        print(f"생성: {path}")


if __name__ == "__main__":
    main()
