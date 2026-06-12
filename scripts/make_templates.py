"""번들 템플릿(templates/*.pptx) 생성 스크립트.

기본 템플릿은 python-pptx 내장 기본 프레젠테이션이고, dark 템플릿은
테마 색상(theme1.xml)을 어두운 배색으로 바꾼 변형이다. 자체 템플릿을
만들 때 참고용으로 사용할 수 있다. 실행::

    python scripts/make_templates.py
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES = ROOT / "templates"

# 어두운 테마 배색: Office 기본 테마의 색상 슬롯을 치환한다
_DARK_COLORS = {
    "dk1": "F5F5F5",      # 본문 텍스트 → 밝게
    "lt1": "1E1E2E",      # 배경 → 어둡게
    "dk2": "E0E0F0",
    "lt2": "2A2A3C",
    "accent1": "89B4FA",
    "accent2": "F38BA8",
    "accent3": "A6E3A1",
    "accent4": "FAB387",
    "accent5": "CBA6F7",
    "accent6": "94E2D5",
}


def make_default() -> Path:
    path = TEMPLATES / "default.pptx"
    Presentation().save(str(path))
    return path


def make_dark(default_path: Path) -> Path:
    path = TEMPLATES / "dark.pptx"
    shutil.copy(default_path, path)

    # pptx는 zip 컨테이너이므로 테마 XML만 치환해서 다시 묶는다
    with zipfile.ZipFile(path) as zf:
        items = {name: zf.read(name) for name in zf.namelist()}

    theme_names = [n for n in items if re.match(r"ppt/theme/theme\d+\.xml", n)]
    for name in theme_names:
        xml = items[name].decode("utf-8")
        for slot, color in _DARK_COLORS.items():
            # <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1> 등
            xml = re.sub(
                rf"(<a:{slot}>).*?(</a:{slot}>)",
                rf'\1<a:srgbClr val="{color}"/>\2',
                xml,
                flags=re.DOTALL,
            )
        items[name] = xml.encode("utf-8")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in items.items():
            zf.writestr(name, data)

    # 손상 여부 검증: python-pptx로 다시 열리는지 확인
    Presentation(str(path))
    return path


def main():
    TEMPLATES.mkdir(exist_ok=True)
    default_path = make_default()
    print(f"생성: {default_path}")
    dark_path = make_dark(default_path)
    print(f"생성: {dark_path}")


if __name__ == "__main__":
    main()
