---
title: AI 코드 망치는 6가지와 해결법
subtitle: “Code is not cheap. Bad code is the most expensive it’s ever been.”
author: Matt Pocock — TypeScript 거장의 AI 시대 소프트웨어 펀더멘털
template: templates/agentos/teal.pptx
---

<!-- layout: section -->
<!-- eyebrow: POCOCK의 핵심 명제 -->
# AI 시대일수록 펀더멘털이 더 중요

분석 깊이가 곧 코드 품질이다

---

<!-- eyebrow: 6 TRAPS · 6 PRESCRIPTIONS -->
# 6가지 함정 + 6가지 처방
- **함정 #1** · AI가 의도와 다른 걸 만든다
- **함정 #2** · 모호한 PRD
- **함정 #3** · 코드 리뷰의 죽음
- **함정 #4** · Shallow Modules
- **함정 #5** · 의존성 방치
- **함정 #6** · 설계의 부재

::: notes
전체 구조 — 각 함정마다 처방이 하나씩 따라온다.
:::

---

<!-- eyebrow: THE PRESCRIPTION -->
# Grill Me — 심문해줘

```grill-me.md
# Grill Me Skill
Interview me relentlessly about
every aspect of this plan,
until we reach a shared understanding.
```

계획의 모든 측면을 — *공유된 이해*에 도달할 때까지 집요하게 인터뷰

---

<!-- layout: two-content -->
<!-- eyebrow: POCOCK'S TAKE -->
# Grill Me vs plan mode

::: left
- **PLAN MODE**
  - 너무 빨리 결과물부터
  - 만들려 함
:::

::: right
- **GRILL ME**
  - 같은 이해 도달이 먼저
  - 공유된 컨텍스트 확보
:::

---

<!-- eyebrow: TRAP #5 -->
# 의존성 관리의 실종

| 함정 | 결과 |
| --- | --- |
| 아무거나 추가 | AI는 라이브러리를 망설이지 않는다 |
| 검토 없는 버전 | 보안 · 라이선스 · 호환성 공백 |
| 공급망 리스크 | npm 한 줄이 시스템 전체를 흔든다 |

---

<!-- layout: section -->
<!-- eyebrow: THE CONCLUSION · KENT BECK -->
# 매일 시스템 설계에 투자하라

— AI는 좋은 코드베이스에서만 빛난다
