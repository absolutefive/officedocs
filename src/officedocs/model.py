"""슬라이드 덱의 중간 표현(IR) 모델.

파서가 입력 텍스트를 이 모델로 변환하고, 빌더가 이 모델과 PPTX 템플릿을
조합해 최종 PPTX를 생성한다. 입력 형식과 출력 템플릿이 서로 독립적이므로
템플릿만 갈아끼우면 같은 내용으로 즉시 재생성할 수 있다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Union


@dataclass
class Run:
    """인라인 서식이 적용된 텍스트 조각."""

    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False


@dataclass
class Paragraph:
    """문단 또는 글머리 기호 항목."""

    runs: List[Run]
    level: int = 0          # 글머리 기호 들여쓰기 단계 (0부터)
    bullet: bool = False    # 글머리 기호 여부
    numbered: bool = False  # 번호 매기기 여부

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)


@dataclass
class Table:
    """표. 첫 행은 머리글로 취급한다."""

    rows: List[List[List[Run]]]  # rows[행][열] = Run 목록
    has_header: bool = True


@dataclass
class Image:
    """이미지 참조. path는 입력 문서 기준 상대 경로를 허용한다."""

    path: str
    alt: str = ""


@dataclass
class CodeBlock:
    """코드 블록. 다크 코드 윈도우(라운드 카드 + 신호등 점 + 파일명)로 렌더링된다."""

    lines: List[str]
    label: str = ""  # 파일명 또는 언어 표시


@dataclass
class Card:
    """카드 한 장: 작은 라벨 + 큰 텍스트 + 보조 설명."""

    big: str = ""
    label: str = ""
    sub: str = ""
    tinted: bool = False  # 포인트 컬러 강조 카드


@dataclass
class CardGrid:
    """라운드 카드 그리드. numbered면 라벨 대신 번호 원형 배지를 단다."""

    cards: List[Card] = field(default_factory=list)
    numbered: bool = False


@dataclass
class Bar:
    """가로로 긴 강조 바(처방 바): 큰 단어 + 작은 라벨 + 본문 한 줄."""

    word: str = ""
    label: str = ""
    text: str = ""


Block = Union[Paragraph, Table, Image, CodeBlock, CardGrid, Bar]


@dataclass
class Slide:
    """슬라이드 한 장."""

    layout: Optional[str] = None     # 시맨틱 레이아웃 이름 (없으면 자동 결정)
    title: Optional[str] = None
    eyebrow: Optional[str] = None    # 제목 위 강조 라벨 (포인트 컬러)
    blocks: List[Block] = field(default_factory=list)
    left: List[Block] = field(default_factory=list)   # two-content 좌측
    right: List[Block] = field(default_factory=list)  # two-content 우측
    notes: Optional[str] = None


@dataclass
class Deck:
    """덱 전체: 메타데이터(머리말)와 슬라이드 목록."""

    meta: Dict[str, str] = field(default_factory=dict)
    slides: List[Slide] = field(default_factory=list)
