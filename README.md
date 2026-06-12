# officedocs

텍스트 문서(마크다운)를 작성하면 **편집 가능한 PPTX**를 자동 생성하는 도구입니다.

- 결과물은 이미지가 아니라 PowerPoint 네이티브 개체(자리표시자, 텍스트 상자, 표, 그림)로 생성되므로 **PowerPoint에서 직접 수정**할 수 있습니다.
- 내용(텍스트 문서)과 디자인(PPTX 템플릿)이 분리되어 있어, **템플릿 파일만 갈아끼우면 같은 내용으로 즉시 재생성**됩니다.
- 핵심 의존성은 [python-pptx](https://github.com/scanny/python-pptx) 하나뿐입니다(MIT 라이선스, 활발히 유지보수되는 사실상 표준 라이브러리). 공급망 리스크를 줄이기 위해 그 외 외부 의존성은 사용하지 않습니다.

## 설치

```bash
pip install -e .
# 개발(테스트 포함):
pip install -e .[dev]
```

Python 3.9 이상이 필요합니다.

## 빠른 시작

```bash
# 1. 입력 문서 스캐폴드 생성
officedocs init my-deck.md

# 2. 내용 작성 후 빌드 (기본 내장 템플릿)
officedocs build my-deck.md -o my-deck.pptx

# 3. 템플릿을 갈아끼워 재생성
officedocs build my-deck.md -t templates/dark.pptx -o my-deck-dark.pptx
officedocs build my-deck.md -t 회사공식템플릿.pptx -o my-deck-corp.pptx
```

전체 예시는 [`examples/sample.md`](examples/sample.md)를 참고하세요.

```bash
officedocs build examples/sample.md -t templates/default.pptx -o sample.pptx
```

## 입력 템플릿 문법

````markdown
---
title: 발표 제목          # 있으면 표지 슬라이드 자동 생성
subtitle: 부제목
author: 작성자
date: 2026-01-01
template: templates/dark.pptx   # 선택: 문서에 템플릿 고정 (-t 옵션이 우선)
---

<!-- layout: section -->
# 섹션 제목

---

# 일반 내용 슬라이드
설명 문단입니다. **굵게**, *기울임*, `코드` 인라인 서식을 지원합니다.

- 글머리 기호
  - 2칸 들여쓰기로 하위 단계
1. 번호 매기기

| 표 | 도 |
| --- | --- |
| 네이티브 | 개체 |

![그림 설명](assets/diagram.png)

::: notes
발표자 노트는 여기에 적습니다.
:::

---

<!-- layout: two-content -->
# 좌우 2단 슬라이드

::: left
- 왼쪽 내용
:::

::: right
- 오른쪽 내용
:::
````

규칙 요약:

| 문법 | 의미 |
| --- | --- |
| 문서 맨 위 `--- ... ---` | 메타데이터(머리말). `title`이 있으면 표지 자동 생성 |
| 단독 줄 `---` | 슬라이드 구분 |
| 슬라이드의 첫 `#`/`##` 제목 | 슬라이드 제목 |
| `<!-- layout: NAME -->` | 레이아웃 지정: `title` `section` `content` `two-content` `title-only` `blank` |
| `- ` / `* ` / `1. ` | 글머리 기호 / 번호 매기기 (2칸 들여쓰기 = 하위 단계) |
| `\| a \| b \|` | PPTX 네이티브 표 (첫 행 + 구분선 = 머리글) |
| `![대체텍스트](경로)` | 그림 개체 (입력 문서 기준 상대 경로) |
| `::: notes` ~ `:::` | 발표자 노트 |
| `::: left` / `::: right` ~ `:::` | 2단 레이아웃의 좌/우 영역 |

레이아웃을 지정하지 않으면 내용 구성에 따라 자동 결정됩니다(좌/우 영역이 있으면 `two-content`, 본문이 있으면 `content` 등).

## 템플릿 교체

어떤 `.pptx` 템플릿이든 `-t` 옵션으로 지정하면 됩니다. 도구가 템플릿의 슬라이드 레이아웃을 다음 순서로 해석합니다.

1. **사이드카 매핑 파일** — 템플릿 옆에 `<템플릿이름>.layouts.json`을 두면 매핑을 직접 제어할 수 있습니다.
   ```json
   {
     "title": "우리회사 표지",
     "content": 3,
     "section": "간지 레이아웃"
   }
   ```
   값은 레이아웃 이름(문자열) 또는 인덱스(숫자)입니다.
2. **레이아웃 이름 키워드** — "Title Slide", "제목 슬라이드", "Section Header" 등 표준 이름 자동 인식
3. **자리표시자 구성 휴리스틱** — 가운데 제목이면 표지, 본문 2개면 2단 등
4. **표준 Office 인덱스** — 최후의 대비책

템플릿의 레이아웃 구성과 자동 매핑 결과는 다음 명령으로 확인합니다.

```bash
officedocs layouts 회사공식템플릿.pptx
```

템플릿 파일에 예시 슬라이드가 남아 있어도 빌드 시 자동 제거되고 레이아웃/마스터/테마만 사용됩니다.

### 프로토타입 방식: 자리표시자가 없는 기업 템플릿

실무 기업 템플릿은 디자인이 슬라이드 레이아웃의 자리표시자가 아니라 **예시 슬라이드의 텍스트 상자**에 들어있는 경우가 많습니다(예: 표지의 "프로젝트 명칭 자리 배치" 같은 안내 텍스트). 이런 템플릿은 사이드카에서 시맨틱 이름에 객체를 주면 **예시 슬라이드를 통째로 복제한 뒤 마커 텍스트를 치환**하는 방식으로 처리합니다.

```json
{
  "title": {
    "prototype": 0,
    "replace": { "프로젝트(솔루션)": "title", "문서(기능)": "subtitle" },
    "remove": ["표지타입", "필요 시 고객사 로고"]
  },
  "content": {
    "prototype": 3,
    "replace": { "속지별 타이틀": "title", "프로젝트(솔루션)": "project" },
    "body": [0.025, 0.13, 0.95, 0.79]
  }
}
```

| 키 | 의미 |
| --- | --- |
| `prototype` | 복제할 템플릿 예시 슬라이드 인덱스 (0부터) |
| `replace` | 마커 부분 문자열 → 채울 필드. 필드: `title` `subtitle` `project`(머리말 title) `author` `date`. 원본 도형의 글꼴/크기/색이 보존됩니다 |
| `remove` | 해당 마커가 포함된 도형 삭제 (안내문구 제거용) |
| `body` | 본문 블록(텍스트/표/그림)을 배치할 영역 `[left, top, width, height]`. 1.0 이하는 슬라이드 크기 대비 비율, 초과는 EMU |

템플릿의 예시 슬라이드는 빌드 결과물에서 자동 제거됩니다. 마커 문자열만 알면 되므로, 새 기업 템플릿을 받았을 때 `officedocs layouts 템플릿.pptx`로 구조를 확인하고 사이드카 JSON 하나만 작성하면 됩니다.

### 번들 템플릿

- `templates/default.pptx`(밝은 기본 테마), `templates/dark.pptx`(어두운 테마) — 레이아웃 방식. `scripts/make_templates.py`로 재생성할 수 있습니다.
- `templates/tomato/typeA-landscape.pptx` 외 4종(TypeA/TypeB × 가로/세로) — 토마토시스템 문서 템플릿. 프로토타입 방식의 실전 예시로, 각각 사이드카 `*.layouts.json`이 함께 있습니다.

```bash
officedocs build examples/sample.md -t templates/tomato/typeA-landscape.pptx -o out.pptx
```

## Python API

```python
from pathlib import Path
from officedocs import parse_document, build_deck

deck = parse_document(Path("deck.md").read_text(encoding="utf-8"))
prs = build_deck(deck, template_path=Path("templates/dark.pptx"), base_dir=Path("."))
prs.save("deck.pptx")
```

## 프로젝트 구조

```
src/officedocs/
  parser.py    # 입력 텍스트 → Deck 모델 (입력 형식 담당)
  model.py     # 중간 표현: Deck / Slide / Paragraph / Table / Image
  layouts.py   # 템플릿 레이아웃 → 시맨틱 이름 해석 (템플릿 교체 담당)
  builder.py   # Deck + 템플릿 → 네이티브 개체로 PPTX 생성
  cli.py       # build / layouts / init 명령
templates/     # 번들 템플릿 (default, dark)
examples/      # 입력 문서 예시
tests/         # pytest 테스트 (파서 + 빌드 라운드트립 검증)
```

## 테스트

```bash
pytest
```

빌드 결과를 메모리에서 저장 후 다시 열어(라운드트립) 표가 네이티브 표 개체인지, 그림이 Picture 개체인지, 본문이 실제 텍스트 프레임인지까지 검증합니다.

## 설계 원칙

- **내용과 디자인의 분리**: 파서는 템플릿을 모르고, 빌더는 입력 문법을 모릅니다. 중간 표현(`model.py`)이 둘을 잇습니다.
- **편집 가능한 출력**: 슬라이드 렌더링 이미지를 붙이는 방식은 쓰지 않습니다. 모든 요소가 PowerPoint 개체입니다.
- **최소 의존성**: 마크다운 파서도 필요한 부분집합만 직접 구현해 외부 의존성을 python-pptx 하나로 유지합니다.
