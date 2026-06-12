/**
 * AgentOS 스타일 PPTX 생성기
 * 영상: "Matt Pocock 41만뷰 발표 — AI 코드 망치는 6가지와 해결법"
 * 스타일: 화이트 배경 / 포인트 컬러 / Pretendard 초고굵기 / 자간 넓은 아이브로우 / 라운드 카드 / 다크 코드 윈도우
 *
 * 사용법: node gen_deck.js   →  4개 컬러웨이 pptx 생성
 * 컬러 교체: PALETTES 배열만 수정하면 전체 덱 색상이 바뀝니다.
 */
const pptxgen = require("pptxgenjs");

// ───────────────────────── 공통 토큰 ─────────────────────────
const INK = "1B1C1E";        // 타이틀 (거의 검정)
const BODY = "3E434B";       // 본문
const MUTED = "8A8F98";      // 보조 회색
const CARD_GRAY = "F4F4F5";  // 중립 카드 배경
const CARD_LINE = "E9E9EC";  // 중립 카드 테두리
const CODE_BG = "1F2335";    // 코드 윈도우 본문
const CODE_HEAD = "171A2B";  // 코드 윈도우 헤더
const CODE_TXT = "E8EAF2";   // 코드 일반 텍스트
const CODE_DIM = "8B90A5";   // 파일명/주석 회색
const DOT_R = "ED6A5E", DOT_Y = "F5BE4F", DOT_G = "61C554";

const F_DISP = "Pretendard ExtraBold"; // 디스플레이 타이틀 (없으면 Bold로 폴백)
const F_MAIN = "Pretendard";           // 본문/라벨
const F_MONO = "D2Coding";             // 코드

const W = 13.3, H = 7.5;

// ───────────────────────── 컬러웨이 4종 ─────────────────────────
const PALETTES = [
  { id: "01_teal_original", name: "Teal (원본)",  ACCENT: "0F8A73", TINT: "EBF6F2", CODE_HL: "9FD8A8" },
  { id: "02_blue",          name: "Blue",         ACCENT: "2563EB", TINT: "EBF1FE", CODE_HL: "9DBDF9" },
  { id: "03_red",           name: "Red",          ACCENT: "DC2626", TINT: "FCEDED", CODE_HL: "F2A6A6" },
  { id: "04_orange",        name: "Orange",       ACCENT: "EA580C", TINT: "FDF0E6", CODE_HL: "F6BE8F" },
];

// ═════════════════════════ 빌더 ═════════════════════════
function buildDeck(P) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
  pres.author = "AgentOS Style Template";
  pres.title = "AI 코드 망치는 6가지와 해결법";

  const newSlide = (notes) => {
    const s = pres.addSlide();
    s.background = { color: "FFFFFF" };
    if (notes) s.addNotes(notes);
    return s;
  };

  // ── 좌상단 헤더: 아이브로우 + 초굵은 타이틀 + 짧은 검정 대시 ──
  const header = (s, eyebrow, title, titleSize = 34) => {
    s.addText(eyebrow, { x: 0.55, y: 0.42, w: 11.5, h: 0.3, margin: 0,
      fontFace: F_MAIN, fontSize: 11.5, bold: true, color: P.ACCENT, charSpacing: 3 });
    s.addText(title, { x: 0.53, y: 0.72, w: 12.2, h: 0.8, margin: 0,
      fontFace: F_DISP, fontSize: titleSize, bold: true, color: INK });
    s.addShape('rect', { x: 0.56, y: 1.68, w: 0.55, h: 0.045, fill: { color: INK } });
  };

  // ── 처방 바 (image 3 하단 스타일) ──
  const rxBar = (s, y, word, label, text, wordW = 1.7) => {
    s.addShape('roundRect', { x: 0.55, y, w: 12.2, h: 0.92, rectRadius: 0.08,
      fill: { color: P.TINT }, line: { color: P.ACCENT, width: 1 } });
    s.addText(word, { x: 0.85, y, w: wordW, h: 0.92, margin: 0, valign: "middle",
      fontFace: F_DISP, fontSize: 20, bold: true, color: P.ACCENT });
    s.addText(label, { x: 0.95 + wordW, y: y + 0.16, w: 9.5, h: 0.26, margin: 0,
      fontFace: F_MAIN, fontSize: 9.5, bold: true, color: P.ACCENT, charSpacing: 2.5 });
    s.addText(text, { x: 0.95 + wordW, y: y + 0.44, w: 10.3, h: 0.32, margin: 0,
      fontFace: F_MAIN, fontSize: 13.5, bold: true, color: INK });
  };

  // ── 비교 카드 한 장 ──
  const compareCard = (s, x, y, w, h, { label, big, sub, tinted }) => {
    s.addShape('roundRect', { x, y, w, h, rectRadius: 0.09,
      fill: { color: tinted ? P.TINT : CARD_GRAY },
      line: tinted ? { color: P.ACCENT, width: 1 } : { color: CARD_LINE, width: 0.75 } });
    s.addText(label, { x, y: y + 0.24, w, h: 0.26, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 10.5, bold: true, color: tinted ? P.ACCENT : MUTED, charSpacing: 2.5 });
    s.addText(big, { x: x + 0.2, y: y + 0.56, w: w - 0.4, h: 0.55, align: "center", margin: 0,
      fontFace: F_DISP, fontSize: 21, bold: true, color: tinted ? P.ACCENT : INK });
    s.addText(sub, { x: x + 0.2, y: y + 1.18, w: w - 0.4, h: 0.32, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 12.5, color: "6B7077" });
  };

  // ── 다크 코드 윈도우 ──
  const codeWindow = (s, x, y, w, h, filename, lines) => {
    s.addShape('roundRect', { x, y, w, h, rectRadius: 0.07, fill: { color: CODE_BG } });
    s.addShape('roundRect', { x, y, w, h: 0.38, rectRadius: 0.07, fill: { color: CODE_HEAD } });
    [DOT_R, DOT_Y, DOT_G].forEach((c, i) =>
      s.addShape('ellipse', { x: x + 0.2 + i * 0.16, y: y + 0.145, w: 0.09, h: 0.09, fill: { color: c } }));
    s.addText(filename, { x: x + 0.75, y: y + 0.07, w: 3, h: 0.25, margin: 0,
      fontFace: F_MONO, fontSize: 10, color: CODE_DIM });
    s.addText(lines.map((ln, i) => ({
      text: ln.text, options: { color: ln.hl ? P.CODE_HL : CODE_TXT, breakLine: i < lines.length - 1 }
    })), { x: x + 0.28, y: y + 0.55, w: w - 0.56, h: h - 0.75, margin: 0, valign: "top",
      fontFace: F_MONO, fontSize: 13, lineSpacing: 23 });
  };

  // ── 번호 원형 배지 ──
  const numBadge = (s, x, y, n, d = 0.3) => {
    s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: P.ACCENT } });
    s.addText(String(n), { x, y, w: d, h: d, align: "center", valign: "middle", margin: 0,
      fontFace: F_MAIN, fontSize: 11.5, bold: true, color: "FFFFFF" });
  };

  // ── ①②③ 리치 라인 ──
  const circledLine = (s, x, y, w, mark, head, rest, size = 15) => {
    s.addText([
      { text: mark + " ", options: { color: P.ACCENT, bold: true } },
      { text: head, options: { color: P.ACCENT, bold: true } },
      { text: rest, options: { color: INK, bold: true } },
    ], { x, y, w, h: 0.4, margin: 0, fontFace: F_MAIN, fontSize: size });
  };

  // ── 센터 인용형 슬라이드 ──
  const quoteSlide = (s, eyebrow, title, sub, caption, titleSize = 42) => {
    s.addText(eyebrow, { x: 0, y: 2.0, w: W, h: 0.3, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 12, bold: true, color: P.ACCENT, charSpacing: 4 });
    s.addText("\u201C", { x: 0, y: 2.34, w: W, h: 0.4, align: "center", margin: 0,
      fontFace: F_DISP, fontSize: 22, bold: true, color: INK });
    s.addText(title, { x: 0.5, y: 2.78, w: W - 1, h: 0.95, align: "center", margin: 0,
      fontFace: F_DISP, fontSize: titleSize, bold: true, color: INK });
    if (sub) s.addText(sub, { x: 0.5, y: 3.95, w: W - 1, h: 0.4, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 15, color: MUTED });
    if (caption) s.addText(caption, { x: 0.5, y: 4.6, w: W - 1, h: 0.4, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 13, bold: true, color: BODY });
  };

  // ═══════════ S1. 커버 ═══════════
  {
    const s = newSlide("인트로 (0:00) — 왜 이 발표가 41만 뷰인가. TypeScript 거장 Matt Pocock의 AI 시대 발표 해설.");
    s.addText("AGENTOS  ·  MATT POCOCK 41만 뷰 발표 해설", { x: 0, y: 2.05, w: W, h: 0.3, align: "center",
      margin: 0, fontFace: F_MAIN, fontSize: 12, bold: true, color: P.ACCENT, charSpacing: 4 });
    s.addText("AI 코드 망치는 6가지와 해결법", { x: 0.5, y: 2.5, w: W - 1, h: 1.0, align: "center", margin: 0,
      fontFace: F_DISP, fontSize: 44, bold: true, color: INK });
    s.addText("\u201CCode is not cheap. Bad code is the most expensive it\u2019s ever been.\u201D",
      { x: 0.5, y: 3.75, w: W - 1, h: 0.4, align: "center", margin: 0, fontFace: F_MAIN, fontSize: 15, color: MUTED });
    s.addText("Matt Pocock  —  TypeScript 거장의 AI 시대 소프트웨어 펀더멘털",
      { x: 0.5, y: 4.45, w: W - 1, h: 0.35, align: "center", margin: 0, fontFace: F_MAIN, fontSize: 13, bold: true, color: BODY });
  }

  // ═══════════ S2. THE PIVOT (image 6) ═══════════
  {
    const s = newSlide("Pocock은 TypeScript 교육자에서 AI 엔지니어링으로 피봇 — AI Hero 플랫폼과 Claude Code for Real Engineers 강의.");
    header(s, "THE PIVOT", "AI 엔지니어링으로 피봇");
    // 작은 카드 2개 + '+'
    s.addShape('roundRect', { x: 0.55, y: 2.25, w: 2.5, h: 1.0, rectRadius: 0.09,
      fill: { color: P.TINT }, line: { color: P.ACCENT, width: 1 } });
    s.addText("플랫폼", { x: 0.8, y: 2.4, w: 2.0, h: 0.22, margin: 0, fontFace: F_MAIN, fontSize: 9.5, bold: true, color: P.ACCENT, charSpacing: 1.5 });
    s.addText("🎓 AI Hero", { x: 0.8, y: 2.66, w: 2.1, h: 0.4, margin: 0, fontFace: F_DISP, fontSize: 16, bold: true, color: P.ACCENT });
    s.addText("+", { x: 3.15, y: 2.55, w: 0.45, h: 0.45, align: "center", valign: "middle", margin: 0, fontFace: F_MAIN, fontSize: 16, color: MUTED });
    s.addShape('roundRect', { x: 3.7, y: 2.25, w: 4.9, h: 1.0, rectRadius: 0.09,
      fill: { color: P.TINT }, line: { color: P.ACCENT, width: 1 } });
    s.addText("강의", { x: 3.95, y: 2.4, w: 2.0, h: 0.22, margin: 0, fontFace: F_MAIN, fontSize: 9.5, bold: true, color: P.ACCENT, charSpacing: 1.5 });
    s.addText("\u201CClaude Code for Real Engineers\u201D", { x: 3.95, y: 2.66, w: 4.5, h: 0.4, margin: 0,
      fontFace: F_DISP, fontSize: 15, bold: true, color: P.ACCENT });
    s.addText("——  이 강의에서 정리한 한 마디  ——", { x: 0.55, y: 3.6, w: 12.2, h: 0.3, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 11, color: MUTED });
    // 큰 결론 카드
    s.addShape('roundRect', { x: 0.55, y: 4.05, w: 12.2, h: 1.75, rectRadius: 0.1,
      fill: { color: P.TINT }, line: { color: P.ACCENT, width: 1 } });
    s.addText("핵심 결론", { x: 0.9, y: 4.28, w: 3, h: 0.24, margin: 0, fontFace: F_MAIN, fontSize: 10, bold: true, color: P.ACCENT, charSpacing: 1.5 });
    s.addText([
      { text: "\u201CAI로 피봇한 사람이 알고 보니", options: { breakLine: true } },
      { text: "옛날 펀더멘털이 더 중요하더라고 결론낸 발표\u201D", options: {} },
    ], { x: 0.9, y: 4.58, w: 11.5, h: 1.0, margin: 0, fontFace: F_DISP, fontSize: 17, bold: true, color: INK, lineSpacing: 28 });
  }

  // ═══════════ S3. THE THESIS (image 7) ═══════════
  {
    const s = newSlide("핵심 명제 (2:00) — AI 시대일수록 소프트웨어 펀더멘털이 더 중요해진다.");
    s.addText("POCOCK의 핵심 명제", { x: 0, y: 1.55, w: W, h: 0.3, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 12, bold: true, color: P.ACCENT, charSpacing: 4 });
    s.addShape('roundRect', { x: 3.35, y: 2.35, w: 6.6, h: 2.7, rectRadius: 0.12,
      fill: { color: P.TINT }, line: { color: P.ACCENT, width: 1.25 } });
    s.addText("THE THESIS", { x: 3.35, y: 2.78, w: 6.6, h: 0.3, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 12, bold: true, color: "0E6B59" === "0E6B59" ? P.ACCENT : P.ACCENT, charSpacing: 3 });
    s.addText([
      { text: "AI 시대일수록", options: { color: INK, breakLine: true } },
      { text: "펀더멘털이 더 중요", options: { color: P.ACCENT } },
    ], { x: 3.35, y: 3.2, w: 6.6, h: 1.5, align: "center", margin: 0, fontFace: F_DISP, fontSize: 27, bold: true, lineSpacing: 42 });
  }

  // ═══════════ S4. Code is not cheap ═══════════
  {
    const s = newSlide("핵심 명제 (2:00) — Code is not cheap. AI는 좋은 코드베이스에서만 빛난다.");
    header(s, "THE CORE CLAIM", "Code is not cheap");
    s.addShape('roundRect', { x: 0.55, y: 2.3, w: 6.3, h: 2.6, rectRadius: 0.1, fill: { color: CODE_BG } });
    s.addText([
      { text: "\u201CBad code is the most expensive", options: { breakLine: true } },
      { text: "it\u2019s ever been.\u201D", options: {} },
    ], { x: 0.95, y: 2.75, w: 5.6, h: 1.2, margin: 0, fontFace: F_DISP, fontSize: 20, bold: true, color: "FFFFFF", lineSpacing: 32 });
    s.addText("— Matt Pocock", { x: 0.95, y: 4.25, w: 4, h: 0.3, margin: 0, fontFace: F_MONO, fontSize: 11, color: CODE_DIM });
    circledLine(s, 7.5, 2.55, 5.4, "\u2460", "좋은 코드베이스", " = AI가 빛나는 곳");
    circledLine(s, 7.5, 3.25, 5.4, "\u2461", "나쁜 코드", " = 역대 가장 비싼 비용");
    circledLine(s, 7.5, 3.95, 5.4, "\u2462", "그래서", " 펀더멘털이 그 어느 때보다 중요");
    s.addText("변경하기 어려운 코드베이스라면 — AI가 줄 수 있는 가치를 누릴 수 없다",
      { x: 0.55, y: 5.45, w: 12, h: 0.35, margin: 0, fontFace: F_MAIN, fontSize: 13, color: MUTED });
  }

  // ═══════════ S5. 6가지 함정 그리드 (image 8) ═══════════
  {
    const s = newSlide("전체 구조 — 6가지 함정과 6가지 처방. 이후 슬라이드에서 하나씩.");
    header(s, "6 TRAPS  ·  6 PRESCRIPTIONS", "6가지 함정 + 6가지 처방");
    const traps = [
      "AI가 의도와 다른 걸 만든다", "모호한 PRD",
      "코드 리뷰의 죽음", "Shallow Modules",
      "의존성 방치", "설계의 부재",
    ];
    traps.forEach((t, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.55 + col * 6.3, y = 2.2 + row * 0.95;
      s.addShape('roundRect', { x, y, w: 5.9, h: 0.68, rectRadius: 0.07,
        fill: { color: CARD_GRAY }, line: { color: CARD_LINE, width: 0.75 } });
      numBadge(s, x + 0.22, y + 0.19, i + 1);
      s.addText([
        { text: `함정 #${i + 1}`, options: { bold: true, color: INK } },
        { text: "  ·  " + t, options: { bold: true, color: BODY } },
      ], { x: x + 0.68, y, w: 5.1, h: 0.68, valign: "middle", margin: 0, fontFace: F_MAIN, fontSize: 13.5 });
    });
    s.addText("각 함정마다 Pocock의 처방(THE PRESCRIPTION)이 하나씩 따라온다",
      { x: 0.55, y: 5.35, w: 12, h: 0.35, margin: 0, fontFace: F_MAIN, fontSize: 12.5, color: MUTED });
  }

  // ═══════════ S6. 함정 #1 ═══════════
  {
    const s = newSlide("함정 #1 (2:36) — AI가 의도와 다른 걸 만든다. 처방: Grill Me 스킬.");
    header(s, "TRAP #1", "AI가 의도와 다른 걸 만든다");
    compareCard(s, 1.0, 2.35, 4.9, 1.7, { label: "내 머릿속", big: "말하지 않은 가정 30개", sub: "의도는 저절로 전달되지 않는다" });
    s.addText("\u2192", { x: 6.05, y: 2.9, w: 0.7, h: 0.6, align: "center", margin: 0, fontFace: F_MAIN, fontSize: 28, color: "9AA0A6" });
    compareCard(s, 6.9, 2.35, 4.9, 1.7, { label: "AI의 결과물", big: "그럴듯하지만 다른 것", sub: "빠르게 · 자신 있게 · 어긋난다", tinted: false });
    rxBar(s, 5.0, "Grill Me", "THE PRESCRIPTION", "코드를 짜기 전에 — AI가 나를 심문하게 하라", 1.85);
  }

  // ═══════════ S7. Grill Me 코드 윈도우 (image 1) ═══════════
  {
    const s = newSlide("처방 #1 — grill-me.md 스킬. 공유된 이해에 도달할 때까지 인터뷰.");
    header(s, "THE PRESCRIPTION", "Grill Me — 심문해줘");
    codeWindow(s, 0.55, 2.25, 9.3, 2.55, "grill-me.md", [
      { text: "# Grill Me Skill", hl: true },
      { text: "Interview me relentlessly about" },
      { text: "every aspect of this plan," },
      { text: "until we reach a shared understanding.", hl: true },
    ]);
    s.addText("계획의 모든 측면을 — 공유된 이해(shared understanding)에 도달할 때까지 집요하게 인터뷰",
      { x: 0.55, y: 5.15, w: 12, h: 0.35, margin: 0, fontFace: F_MAIN, fontSize: 13, color: MUTED });
  }

  // ═══════════ S8. POCOCK'S TAKE (image 2) ═══════════
  {
    const s = newSlide("Pocock의 평가 — Grill Me가 Claude Code plan mode보다 낫다. plan mode는 결과물부터 만들려 하고, Grill Me는 같은 이해 도달이 먼저.");
    s.addText("POCOCK\u2019S TAKE", { x: 0, y: 0.85, w: W, h: 0.3, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 12, bold: true, color: P.ACCENT, charSpacing: 4 });
    s.addText("Grill Me  >  Claude Code plan mode", { x: 0.5, y: 1.2, w: W - 1, h: 0.7, align: "center", margin: 0,
      fontFace: F_DISP, fontSize: 30, bold: true, color: INK });
    compareCard(s, 1.0, 2.5, 4.9, 1.7, { label: "PLAN MODE", big: "⚡ 너무 빨리", sub: "결과물부터 만들려 함" });
    s.addText("\u2192", { x: 6.05, y: 3.05, w: 0.7, h: 0.6, align: "center", margin: 0, fontFace: F_MAIN, fontSize: 28, color: "9AA0A6" });
    compareCard(s, 6.9, 2.5, 4.9, 1.7, { label: "GRILL ME", big: "🤝 같은 이해", sub: "도달이 먼저", tinted: true });
    s.addText("🎬 본인 입으로 들어볼게요 \u2192", { x: 0.5, y: 4.75, w: W - 1, h: 0.35, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 13, bold: true, color: BODY });
  }

  // ═══════════ S9. 함정 #2 ═══════════
  {
    const s = newSlide("함정 #2 (5:05) — 모호한 PRD. 한 줄 요구사항의 빈칸은 AI가 멋대로 채운다. 처방: 스펙 먼저.");
    header(s, "TRAP #2", "모호한 PRD");
    compareCard(s, 1.0, 2.35, 4.9, 1.7, { label: "BEFORE", big: "\u201C로그인 만들어줘\u201D", sub: "빈칸은 AI가 멋대로 채운다" });
    s.addText("\u2192", { x: 6.05, y: 2.9, w: 0.7, h: 0.6, align: "center", margin: 0, fontFace: F_MAIN, fontSize: 28, color: "9AA0A6" });
    compareCard(s, 6.9, 2.35, 4.9, 1.7, { label: "AFTER", big: "행동 단위 스펙", sub: "입력 · 출력 · 엣지케이스 먼저", tinted: true });
    rxBar(s, 5.0, "SPEC", "THE PRESCRIPTION", "모호함은 AI 시대에 가장 비싸게 갚는 빚 — 합의된 스펙부터 쓴다", 1.3);
  }

  // ═══════════ S10. 함정 #3 ═══════════
  {
    const s = newSlide("함정 #3 (5:05) — 코드 리뷰의 죽음. 생성 속도가 리뷰 속도를 추월. 처방: 작은 diff와 인터페이스 중심 리뷰.");
    header(s, "TRAP #3", "코드 리뷰의 죽음");
    s.addShape('roundRect', { x: 0.55, y: 2.3, w: 5.7, h: 2.25, rectRadius: 0.1, fill: { color: CODE_BG } });
    s.addText("\u201CLGTM\u201D — 읽지 않고 머지", { x: 0.95, y: 2.85, w: 5.0, h: 0.5, margin: 0,
      fontFace: F_DISP, fontSize: 20, bold: true, color: "FFFFFF" });
    s.addText("생성 속도가 리뷰 속도를 추월했다", { x: 0.95, y: 3.5, w: 4.8, h: 0.3, margin: 0,
      fontFace: F_MONO, fontSize: 11.5, color: CODE_DIM });
    circledLine(s, 6.8, 2.55, 6.0, "\u2460", "안 읽은 코드", " = 자산이 아니라 부채");
    circledLine(s, 6.8, 3.25, 6.0, "\u2461", "거대한 diff", " = 리뷰 불가능");
    circledLine(s, 6.8, 3.95, 6.0, "\u2462", "리뷰 단위", "를 다시 설계하라");
    rxBar(s, 5.0, "REVIEW", "THE PRESCRIPTION", "작은 diff · 인터페이스 중심 — 리뷰를 부활시켜라", 1.7);
  }

  // ═══════════ S11. THE METAPHOR (image 3) ═══════════
  {
    const s = newSlide("메타포 — Outrunning Your Headlights (Pragmatic Programmer). 피드백의 속도 = 너의 속도 제한. 처방: TDD.");
    header(s, "THE METAPHOR", "\u201COutrunning Your Headlights\u201D");
    s.addText("Pragmatic Programmer 인용", { x: 0.55, y: 1.95, w: 6, h: 0.3, margin: 0,
      fontFace: F_MAIN, fontSize: 12.5, color: MUTED });
    // 자동차 일러스트
    s.addShape('line', { x: 0.9, y: 4.22, w: 3.4, h: 0, line: { color: INK, width: 1.5 } });
    s.addShape('roundRect', { x: 1.1, y: 3.6, w: 0.9, h: 0.45, rectRadius: 0.08, fill: { color: INK } });
    s.addShape('ellipse', { x: 1.22, y: 4.0, w: 0.19, h: 0.19, fill: { color: INK } });
    s.addShape('ellipse', { x: 1.66, y: 4.0, w: 0.19, h: 0.19, fill: { color: INK } });
    s.addShape('triangle', { x: 2.18, y: 3.45, w: 0.62, h: 0.92, rotate: 90, fill: { color: "FAE9A8" } });
    s.addShape("star8", { x: 3.15, y: 3.42, w: 0.52, h: 0.52, fill: { color: "F4A93C" }, line: { color: "E2574C", width: 1 } });
    // 핵심 문장
    s.addText("피드백의 속도  =  너의 속도 제한", { x: 5.3, y: 3.55, w: 7.4, h: 0.6, margin: 0,
      fontFace: F_DISP, fontSize: 23, bold: true, color: INK });
    rxBar(s, 5.35, "TDD", "THE PRESCRIPTION", "테스트 먼저 \u2192 AI가 강제로 작은 단위로 — 헤드라이트 추월 X", 1.1);
  }

  // ═══════════ S12. 함정 #4 Deep vs Shallow (도식) ═══════════
  {
    const s = newSlide("함정 #4 (7:27) — Shallow Modules. Deep module = 좁은 인터페이스 + 깊은 구현 (Ousterhout).");
    header(s, "TRAP #4", "Deep vs Shallow Modules");
    // DEEP (좋음, tinted)
    s.addShape('roundRect', { x: 0.55, y: 2.25, w: 5.9, h: 3.0, rectRadius: 0.1,
      fill: { color: P.TINT }, line: { color: P.ACCENT, width: 1 } });
    s.addText("DEEP MODULE  ·  ✓", { x: 0.55, y: 2.48, w: 5.9, h: 0.26, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 11, bold: true, color: P.ACCENT, charSpacing: 2.5 });
    s.addShape('rect', { x: 2.85, y: 2.92, w: 1.3, h: 0.14, fill: { color: P.ACCENT } });
    s.addShape('rect', { x: 2.85, y: 3.06, w: 1.3, h: 1.45, fill: { color: "D6D9DE" } });
    s.addText("좁은 인터페이스 · 깊은 구현 — AI에게 통째로 맡기기 좋다",
      { x: 0.85, y: 4.7, w: 5.3, h: 0.35, align: "center", margin: 0, fontFace: F_MAIN, fontSize: 12.5, bold: true, color: INK });
    // SHALLOW (나쁨, neutral)
    s.addShape('roundRect', { x: 6.85, y: 2.25, w: 5.9, h: 3.0, rectRadius: 0.1,
      fill: { color: CARD_GRAY }, line: { color: CARD_LINE, width: 0.75 } });
    s.addText("SHALLOW MODULE  ·  ✕", { x: 6.85, y: 2.48, w: 5.9, h: 0.26, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 11, bold: true, color: MUTED, charSpacing: 2.5 });
    s.addShape('rect', { x: 7.95, y: 3.42, w: 3.7, h: 0.14, fill: { color: "9AA0A6" } });
    s.addShape('rect', { x: 7.95, y: 3.56, w: 3.7, h: 0.4, fill: { color: "D6D9DE" } });
    s.addText("넓은 인터페이스 · 얕은 구현 — 복잡도가 밖으로 샌다",
      { x: 7.15, y: 4.7, w: 5.3, h: 0.35, align: "center", margin: 0, fontFace: F_MAIN, fontSize: 12.5, bold: true, color: BODY });
    s.addText("John Ousterhout — A Philosophy of Software Design",
      { x: 0.55, y: 5.5, w: 12, h: 0.3, margin: 0, fontFace: F_MAIN, fontSize: 11.5, color: MUTED });
    rxBar(s, 6.0, "DEPTH", "THE PRESCRIPTION", "AI 위임의 단위는 함수가 아니라 — 깊은 모듈", 1.5);
  }

  // ═══════════ S13. Gray Box 전략 (image 4) ═══════════
  {
    const s = newSlide("처방 #4 — Gray Box 전략. 인터페이스는 사람이 설계, 구현은 AI에게 통째로 위임, 검증은 인터페이스 단위로만.");
    header(s, "THE PRESCRIPTION", "Gray Box 전략");
    s.addText("👤 INTERFACE (사람이 설계)", { x: 1.15, y: 2.5, w: 3.1, h: 0.26, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 10.5, bold: true, color: P.ACCENT, charSpacing: 1 });
    s.addShape('rect', { x: 1.15, y: 2.82, w: 3.1, h: 0.035, fill: { color: P.ACCENT } });
    s.addShape('roundRect', { x: 1.15, y: 3.0, w: 3.1, h: 1.75, rectRadius: 0.1, fill: { color: "9CA2AA" } });
    s.addText("GRAY BOX", { x: 1.15, y: 3.5, w: 3.1, h: 0.4, align: "center", margin: 0,
      fontFace: F_DISP, fontSize: 17, bold: true, color: "FFFFFF" });
    s.addText("안쪽 구현 \u2192 AI한테 통째로 위임", { x: 1.15, y: 3.98, w: 3.1, h: 0.3, align: "center", margin: 0,
      fontFace: F_MAIN, fontSize: 10.5, color: "FFFFFF" });
    circledLine(s, 5.6, 2.85, 7.0, "\u2460", "인터페이스", " = 사람이 직접 설계");
    circledLine(s, 5.6, 3.55, 7.0, "\u2461", "구현", " = AI한테 통째로 위임");
    circledLine(s, 5.6, 4.25, 7.0, "\u2462", "검증", " = 인터페이스 단위로만");
  }

  // ═══════════ S14. 함정 #5 ═══════════
  {
    const s = newSlide("함정 #5 (9:26) — 의존성 방치. AI는 라이브러리 추가를 망설이지 않는다. 처방: 의존성 게이트.");
    header(s, "TRAP #5", "의존성 관리의 실종");
    const cards = [
      ["아무거나 추가", "AI는 라이브러리를 망설이지 않는다"],
      ["검토 없는 버전", "보안 · 라이선스 · 호환성 공백"],
      ["공급망 리스크", "npm 한 줄이 시스템 전체를 흔든다"],
    ];
    cards.forEach((c, i) => {
      const x = 0.55 + i * 4.17;
      s.addShape('roundRect', { x, y: 2.35, w: 3.85, h: 1.65, rectRadius: 0.09,
        fill: { color: CARD_GRAY }, line: { color: CARD_LINE, width: 0.75 } });
      numBadge(s, x + 0.28, y0 = 2.62, i + 1);
      s.addText(c[0], { x: x + 0.72, y: 2.58, w: 3.0, h: 0.4, margin: 0, fontFace: F_DISP, fontSize: 16, bold: true, color: INK });
      s.addText(c[1], { x: x + 0.3, y: 3.18, w: 3.3, h: 0.6, margin: 0, fontFace: F_MAIN, fontSize: 12, color: "6B7077" });
    });
    rxBar(s, 4.65, "GATE", "THE PRESCRIPTION", "의존성 추가는 사람이 승인한다 — AI는 제안까지만", 1.35);
  }

  // ═══════════ S15. 함정 #6 ═══════════
  {
    const s = newSlide("함정 #6 (9:26) — 설계의 부재. 프롬프트는 쌓이는데 설계가 없다. 처방: 매일 설계에 투자 (Kent Beck).");
    header(s, "TRAP #6", "설계 없는 누적");
    circledLine(s, 0.6, 2.5, 6.0, "\u2460", "프롬프트", "는 쌓이는데 설계는 없다");
    circledLine(s, 0.6, 3.2, 6.0, "\u2461", "변경 비용", "이 매주 비싸진다");
    circledLine(s, 0.6, 3.9, 6.0, "\u2462", "어느 날", " AI도 길을 잃는다");
    s.addShape('roundRect', { x: 7.0, y: 2.35, w: 5.75, h: 2.2, rectRadius: 0.1, fill: { color: CODE_BG } });
    s.addText([
      { text: "\u201CInvest in the design of", options: { breakLine: true } },
      { text: "the system every day.\u201D", options: {} },
    ], { x: 7.4, y: 2.75, w: 5.0, h: 0.95, margin: 0, fontFace: F_DISP, fontSize: 18, bold: true, color: "FFFFFF", lineSpacing: 29 });
    s.addText("— Kent Beck", { x: 7.4, y: 3.95, w: 3, h: 0.3, margin: 0, fontFace: F_MONO, fontSize: 11, color: CODE_DIM });
    rxBar(s, 5.0, "DESIGN", "THE PRESCRIPTION", "설계는 이벤트가 아니라 — 매일 하는 투자다", 1.6);
  }

  // ═══════════ S16. 결론 (Kent Beck) ═══════════
  {
    const s = newSlide("결론 (10:36) — Kent Beck: 매일 시스템 설계에 투자하라.");
    quoteSlide(s, "THE CONCLUSION  ·  KENT BECK",
      "매일 시스템 설계에 투자하라",
      "— AI는 좋은 코드베이스에서만 빛난다",
      "\u201CCode is not cheap. Bad code is the most expensive it\u2019s ever been.\u201D");
  }

  // ═══════════ S17. 클로징 (image 5) ═══════════
  {
    const s = newSlide("클로징 — 이 채널의 메시지: 20년 펀더멘털이 너의 무기다.");
    quoteSlide(s, "이  채널의  메시지",
      "20년 펀더멘털이 너의 무기다",
      "—  새로운 도구가 아니라",
      "AI 시대 생존 가이드 — 핵심은 옛날 펀더멘털");
  }

  return pres;
}

// ═════════════════════════ 실행 ═════════════════════════
(async () => {
  for (const P of PALETTES) {
    const pres = buildDeck(P);
    const fn = `/home/claude/AI코드_6가지함정_${P.id}.pptx`;
    await pres.writeFile({ fileName: fn });
    console.log("WROTE", fn);
  }
})();
