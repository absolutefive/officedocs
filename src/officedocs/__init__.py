"""officedocs: 텍스트 문서 → 편집 가능한 PPTX 자동 생성."""

from officedocs.builder import build_deck
from officedocs.parser import parse_document

__version__ = "0.1.0"

__all__ = ["parse_document", "build_deck", "__version__"]
