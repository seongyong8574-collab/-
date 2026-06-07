# -*- coding: utf-8 -*-
"""
moét X ONVIEW 팝업스토어 분석 보고서 — Apple 디자인 시스템 적용본.
풀블리드 타일 교차(화이트/패치먼트 ↔ 니어블랙), 단일 Action Blue 액센트,
SF Pro 느낌(Pretendard 대체)의 타이트한 weight 600 헤드라인,
제품 이미지에만 들어가는 단 하나의 드롭섀도우.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- design tokens
C = {
    "primary":        "0066CC",  # Action Blue
    "primary_focus":  "0071E3",
    "primary_on_dark":"2997FF",  # Sky Link Blue
    "canvas":         "FFFFFF",
    "parchment":      "F5F5F7",
    "pearl":          "FAFAFC",
    "tile1":          "272729",
    "tile2":          "2A2A2C",
    "tile3":          "252527",
    "black":          "000000",
    "ink":            "1D1D1F",
    "on_dark":        "FFFFFF",
    "body_muted":     "CCCCCC",
    "ink80":          "333333",
    "ink48":          "7A7A7A",
    "hairline":       "E0E0E0",
    "divider_soft":   "F0F0F0",
}

# SF Pro substitute optimized for Korean; PowerPoint substitutes if absent.
F_D = "Pretendard"   # SF Pro Display
F_T = "Pretendard"   # SF Pro Text

def rgb(h): return RGBColor.from_string(h)

EMU_IN = 914400
SW, SH = 13.333, 7.5

prs = Presentation()
prs.slide_width  = Emu(int(SW * EMU_IN))
prs.slide_height = Emu(int(SH * EMU_IN))
BLANK = prs.slide_layouts[6]


# ----------------------------------------------------------------- primitives
def slide(bg):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    r.fill.solid(); r.fill.fore_color.rgb = rgb(bg)
    r.line.fill.background(); r.shadow.inherit = False
    return s

def _radius(sp, frac):
    try: sp.adjustments[0] = frac
    except Exception: pass

def rrect(s, x, y, w, h, fill=None, lc=None, lw=1.0, radius=0.06):
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                            Inches(x), Inches(y), Inches(w), Inches(h))
    _radius(sp, radius)
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    if lc is None: sp.line.fill.background()
    else: sp.line.color.rgb = rgb(lc); sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    return sp

def rect(s, x, y, w, h, fill):
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                            Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    sp.line.fill.background(); sp.shadow.inherit = False
    return sp

def line(s, x, y, w, color=C["hairline"], weight=1.0):
    sp = s.shapes.add_connector(2, Inches(x), Inches(y), Inches(x + w), Inches(y))
    sp.line.color.rgb = rgb(color); sp.line.width = Pt(weight)
    sp.shadow.inherit = False
    return sp

def product_shadow(sp):
    """The single allowed drop-shadow: rgba(0,0,0,0.22) 3px 5px 30px — product imagery only."""
    spPr = sp._element.spPr
    el = spPr.makeelement(qn('a:effectLst'), {})
    sh = el.makeelement(qn('a:outerShdw'),
                        {'blurRad': '285750', 'dist': '55539',
                         'dir': '3542400', 'rotWithShape': '0'})
    clr = sh.makeelement(qn('a:srgbClr'), {'val': '000000'})
    clr.append(clr.makeelement(qn('a:alpha'), {'val': '22000'}))
    sh.append(clr); el.append(sh); spPr.append(el)

def _spc(run, px): run.font._rPr.set('spc', str(int(px * 100)))

def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        wrap=True, sp_after=0, ls=None):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE; tf.vertical_anchor = anchor
    for m in ("left", "right", "top", "bottom"): setattr(tf, "margin_" + m, 0)
    first = True
    for text, font, size, color, bold, tracking in runs:
        for seg in text.split("\n"):
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = align
            if sp_after: p.space_after = Pt(sp_after)
            if ls: p.line_spacing = ls
            r = p.add_run(); r.text = seg
            r.font.name = font; r.font.size = Pt(size); r.font.bold = bold
            r.font.color.rgb = rgb(color)
            rPr = r.font._rPr
            for tag in ('a:latin', 'a:ea', 'a:cs'):
                e = rPr.find(qn(tag))
                if e is None: e = rPr.makeelement(qn(tag), {}); rPr.append(e)
                e.set('typeface', font)
            if tracking: _spc(r, tracking)
    return tb

def pill(s, x, y, label, bg=C["primary"], fg=C["on_dark"], w=1.9, h=0.5, outline=None):
    rrect(s, x, y, w, h, fill=bg, lc=outline, lw=1.0, radius=0.5)
    txt(s, x, y, w, h, [(label, F_T, 13, fg, False, -0.2)],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return w, h

def photo(s, x, y, w, h, caption, dark=False):
    base = C["tile3"] if dark else C["parchment"]
    fg   = C["body_muted"] if dark else C["ink48"]
    sp = rrect(s, x, y, w, h, fill=base, radius=0.045)
    product_shadow(sp)  # the one true shadow — product imagery
    gw = 0.55
    rrect(s, x + w/2 - gw/2, y + h/2 - 0.46, gw, 0.36,
          fill=(C["tile2"] if dark else C["hairline"]), radius=0.16)
    txt(s, x + 0.25, y + h/2 + 0.02, w - 0.5, 0.7,
        [("제품/현장 이미지", F_T, 9.5, fg, False, -0.2),
         ("\n" + caption, F_T, 9.5, fg, False, -0.2)],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, ls=1.15)

def eyebrow(s, x, y, text, dark=False):
    col = C["primary_on_dark"] if dark else C["primary"]
    txt(s, x, y, 9.0, 0.34, [(text, F_T, 13, col, True, -0.2)])

def footer(s, n, dark=False):
    fg = C["body_muted"] if dark else C["ink48"]
    txt(s, 0.95, SH - 0.52, 7.0, 0.3,
        [("moét × ONVIEW · BEAUTY PLAYGROUND IN SEONGSU", F_T, 9.5, fg, False, -0.12)],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, SW - 2.0, SH - 0.52, 1.05, 0.3,
        [(f"{n:02d} / 09", F_T, 9.5, fg, False, -0.12)],
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

def header(s, num, kor, eng, subtitle, dark=False):
    ink = C["on_dark"] if dark else C["ink"]
    sub = C["body_muted"] if dark else C["ink48"]
    eyebrow(s, 0.95, 0.72, f"{num}  {eng.upper()}", dark=dark)
    txt(s, 0.93, 1.12, 11.4, 0.85, [(kor, F_D, 36, ink, True, -0.6)])
    txt(s, 0.95, 1.92, 11.4, 0.5, [(subtitle, F_T, 15, sub, False, -0.37)], ls=1.2)

def ucard(s, x, y, w, h, head, body, dark=False, accent=False):
    if dark:
        rrect(s, x, y, w, h, fill=C["tile2"], radius=0.07)
        hc, bc = C["on_dark"], C["body_muted"]
    else:
        rrect(s, x, y, w, h, fill=C["canvas"], lc=C["hairline"], lw=1.0, radius=0.07)
        hc, bc = C["ink"], C["ink80"]
    pad = 0.34
    if accent:
        col = C["primary_on_dark"] if dark else C["primary"]
        rrect(s, x + pad, y + 0.3, 0.42, 0.1, fill=col, radius=0.5)
        ho = 0.5
    else:
        ho = 0.3
    txt(s, x + pad, y + ho, w - 2*pad, 0.5, [(head, F_T, 14.5, hc, True, -0.37)])
    txt(s, x + pad, y + ho + 0.46, w - 2*pad, h - ho - 0.7,
        [(body, F_T, 11.5, bc, False, -0.24)], ls=1.32)


# =====================================================================  SLIDE 1  (light hero)
s = slide(C["parchment"])
eyebrow(s, 0, 1.5, "팝업 리테일 공간 분석 보고서")
# center the eyebrow + headline stack
txt(s, 0.5, 1.45, SW - 1.0, 0.4, [("팝업 리테일 공간 분석 보고서", F_T, 15, C["primary"], True, -0.2)],
    align=PP_ALIGN.CENTER)
txt(s, 0.5, 2.05, SW - 1.0, 2.4,
    [("BEAUTY PLAYGROUND\nIN SEONGSU", F_D, 60, C["ink"], True, -1.4)],
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, ls=1.02)
txt(s, 0.5, 4.55, SW - 1.0, 0.9,
    [("리테일 공간 팝업스토어 분석 보고서", F_T, 21, C["ink80"], False, -0.2),
     ("\n모에뜨(moét) 브랜드존 중심  ·  moét × ONVIEW", F_T, 16, C["ink48"], False, -0.2)],
    align=PP_ALIGN.CENTER, ls=1.35)
# two pill CTAs (Apple double-pill grammar)
pill(s, SW/2 - 2.05, 5.75, "분석 개요", bg=C["primary"], fg=C["on_dark"], w=1.9, h=0.5)
pill(s, SW/2 + 0.15, 5.75, "종합 의견", bg=None, fg=C["primary"], w=1.9, h=0.5, outline=C["primary"])
line(s, SW/2 - 3.0, 6.7, 6.0, color=C["hairline"])
txt(s, 0.5, 6.85, SW - 1.0, 0.4,
    [("14주차 과제 레포트   ·   과목명 _______   ·   학번 _______   ·   이름 _______",
      F_T, 12, C["ink48"], False, -0.12)], align=PP_ALIGN.CENTER)

# =====================================================================  SLIDE 2  (white)
s = slide(C["canvas"])
eyebrow(s, 0.95, 0.72, "OVERVIEW")
txt(s, 0.93, 1.12, 11.4, 0.8, [("개요", F_D, 36, C["ink"], True, -0.6)])
line(s, 0.95, 2.0, SW - 1.9)
txt(s, 0.95, 2.28, 6.2, 0.4, [("WHAT", F_T, 12, C["primary"], True, -0.12)])
txt(s, 0.95, 2.66, 6.2, 3.4,
    [("성수동 중심가에서 개최된 인디 뷰티 큐레이션 플랫폼 ‘오엔뷰(ONVIEW)’와 모에뜨의 "
      "협업 팝업스토어. K-뷰티 고관여 MZ세대를 타깃으로, 일시적 제품 판매를 넘어 "
      "오감 체험(시향·제형)과 팝업 한정 리워드를 제공하는 ‘경험형 리테일 매장’으로 기획되었음.",
      F_T, 16, C["ink80"], False, -0.37)], ls=1.5)
# early-close notice (pearl capsule feel)
rrect(s, 0.95, 5.5, 6.2, 0.95, fill=C["pearl"], lc=C["hairline"], lw=1.0, radius=0.10)
txt(s, 1.25, 5.66, 5.7, 0.7,
    [("운영 유의", F_T, 12, C["primary"], True, -0.2),
     ("\n모에뜨 브랜드존은 6월 4일(목)까지만 조기 운영", F_T, 13, C["ink"], False, -0.37)], ls=1.3)
# info utility card
tx, ty, tw = 7.55, 2.28, 4.85
rrect(s, tx, ty, tw, 4.17, fill=C["canvas"], lc=C["hairline"], lw=1.0, radius=0.06)
rows = [
    ("브랜드", "모에뜨 (moét)"),
    ("팝업스토어명", "BEAUTY PLAYGROUND IN SEONGSU"),
    ("행사 기간", "2026.05.22(금) ~ 06.10(수)"),
    ("운영 시간", "11:00 ~ 19:00 (입장 마감 18:30)"),
    ("장소", "맵달SEOUL 성수 3F\n성동구 성수이로16길 5"),
]
rh = 4.17 / len(rows)
for i, (k, v) in enumerate(rows):
    ry = ty + i * rh
    if i: line(s, tx + 0.32, ry, tw - 0.64, color=C["divider_soft"])
    txt(s, tx + 0.34, ry, 1.55, rh, [(k, F_T, 11, C["ink48"], False, -0.24)],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, tx + 1.85, ry, tw - 2.1, rh, [(v, F_T, 12, C["ink"], True, -0.24)],
        anchor=MSO_ANCHOR.MIDDLE, ls=1.15)
footer(s, 2)

# =====================================================================  SLIDE 3  (dark tile)
s = slide(C["tile1"])
header(s, "01", "장소성", "Locality",
       "트렌드세터의 중심지, 성수동 상권과의 유기적 결합", dark=True)
cards = [
    ("트렌디한 소비층 밀집",
     "국내외 MZ세대 및 K-뷰티 고관여 소비자가 가장 활발하게 유입되는 서울 성수동 중심 상권에 위치."),
    ("우수한 접근성",
     "성수역 3번 출구에서 도보 5분 거리에 위치해 자연스러운 도보 방문객 유입률 확보."),
    ("공간의 적합성",
     "‘맵달서울’ 3층을 활용, 대형 채널에서 못 보던 신선한 인디 뷰티를 발굴·체험하는 성수동 특유의 감성과 시너지."),
]
cw = (SW - 1.9 - 2*0.28) / 3
for i, (h, b) in enumerate(cards):
    ucard(s, 0.95 + i * (cw + 0.28), 2.55, cw, 2.3, h, b, dark=True, accent=(i == 0))
photo(s, 0.95, 5.1, SW - 1.9, 1.75, "3층 행사장 진입 복도 및 입구 대기 줄", dark=True)
footer(s, 3, dark=True)

# =====================================================================  SLIDE 4  (white, 2x2)
s = slide(C["canvas"])
header(s, "02", "상징성", "Symbolism",
       "상품 · 컬러 · 컨셉 · 브랜드 메시지로 구축한 브랜드 아이덴티티")
quad = [
    ("[1] 상품  Products",
     "시그니처 스칼프 샴푸(톳수·창포뿌리줄기수 베이스, 식약처 허가 최대치 식물성 카페인)와 두피 앰플"
     "(마이크로 스피큘 니들샷 공법)을 중앙 매대에 배치, 고기능성 기술력을 직관적으로 체험.", False),
    ("[2] 컬러  Color",
     "메인 스카이 블루(청량함·두피 회복)로 시선을 사로잡고, 미니멀 화이트&그레이(정직함·과학적 근거)로 "
     "패키지·베이스 톤을 구성해 신뢰감 있는 무드 조성.", False),
    ("[3] 컨셉  Concept",
     "원료 설명 카드·감성 엽서로 ‘정직한 성분 연구소’를 전달. 스크런치 안내와 티코스터 증정을 통해 "
     "‘산뜻한 데일리 루틴’을 경험으로 소유하게 설계.", False),
    ("[4] 브랜드 메시지  Message",
     "“Pure Origin, Real Change” — 순수한 핵심 원료(Pure Origin)를 고함량 설계해 실제 두피 개선과 "
     "긍정적 효과(Real Change)를 체감하게 한다는 약속.", True),
]
cw = (SW - 1.9 - 0.28) / 2
ch = 1.92
for i, (h, b, acc) in enumerate(quad):
    cx = 0.95 + (i % 2) * (cw + 0.28)
    cy = 2.5 + (i // 2) * (ch + 0.26)
    ucard(s, cx, cy, cw, ch, h, b, dark=False, accent=acc)
footer(s, 4)

# =====================================================================  SLIDE 5  (dark tile)
s = slide(C["tile1"])
header(s, "03", "테마성", "Thematic",
       "‘Pure Origin, Real Change’ — 체험 중심의 고기능성 헤어 랩(Lab)", dark=True)
cards = [
    ("메인 테마의 구체화",
     "‘뷰티 놀이터(Playground)’라는 전체 테마 속에서 ‘자극 없는 일상 속 두피 회복 솔루션’을 서브 테마로 명확히 제안."),
    ("콘셉추얼한 매대 구성",
     "단순 진열·판매를 넘어, 방문객이 자신의 두피 고민(열감·모발 빠짐·유분)을 인지하고 해답을 스스로 탐색하는 서사적 공간으로 연출."),
]
cw = (SW - 1.9 - 0.28) / 2
for i, (h, b) in enumerate(cards):
    ucard(s, 0.95 + i * (cw + 0.28), 2.55, cw, 2.3, h, b, dark=True, accent=(i == 0))
photo(s, 0.95, 5.1, SW - 1.9, 1.75, "원료 설명 리플렛·감성 엽서가 함께 디스플레이된 매대", dark=True)
footer(s, 5, dark=True)

# =====================================================================  SLIDE 6  (white)
s = slide(C["canvas"])
header(s, "04", "콘텐츠", "Contents",
       "방문객의 오감을 자극하는 참여형 저니(Journey) 설계")
cards = [
    ("원료 스토리텔링 · 시향",
     "특허 ‘톳인삼농축액’, 고함량 식물성 카페인을 코칭하고, 세련된 플로럴 머스크 향을 직접 맡게 하는 청각·후각 콘텐츠."),
    ("촉각적 제형 테스트",
     "마이크로 스피큘(니들샷) 두피 앰플을 손등에 발라보게 해 ‘끈적임 없는 흡수력’과 ‘즉각적 쿨링감’을 현장 체감."),
    ("엔터테인먼트 요소",
     "카카오톡 채널 추가·인스타그램 팔로우 시 ‘꽝 없는 100% 당첨 럭키드로우’ 뽑기 게임으로 참여 재미 극대화."),
]
cw = (SW - 1.9 - 2*0.28) / 3
for i, (h, b) in enumerate(cards):
    ucard(s, 0.95 + i * (cw + 0.28), 2.55, cw, 2.3, h, b, dark=False, accent=(i == 1))
photo(s, 0.95, 5.1, SW - 1.9, 1.75, "두피 앰플 테스팅 장면 및 럭키드로우 돌리기 기계")
footer(s, 6)

# =====================================================================  SLIDE 7  (dark tile)
s = slide(C["tile1"])
header(s, "05", "화제성", "Buzz",
       "‘혜자 팝업’ 입소문과 SNS 자발적 바이럴 효과", dark=True)
cards = [
    ("얼리버드 유입 (오픈런)",
     "“다양한 인디 뷰티 본품 혜택과 다채로운 이벤트”라는 정보가 온라인 뷰티 커뮤니티·SNS로 확산되며 오전 11시 오픈 전부터 대기 수요 발생."),
    ("공유 가치가 있는 굿즈",
     "소장 가치 높은 티코스터와 세련된 스크런치 증정으로, 방문객이 자발적으로 #모에뜨 #성수팝업 인증샷을 업로드하도록 유도."),
]
cw = (SW - 1.9 - 0.28) / 2
for i, (h, b) in enumerate(cards):
    ucard(s, 0.95 + i * (cw + 0.28), 2.55, cw, 2.3, h, b, dark=True, accent=(i == 0))
photo(s, 0.95, 5.1, SW - 1.9, 1.75, "인스타그램 피드 캡처 · 방문 인플루언서 인증샷 레이아웃", dark=True)
footer(s, 7, dark=True)

# =====================================================================  SLIDE 8  (white + stat tile)
s = slide(C["canvas"])
header(s, "06", "유도성", "Inducement",
       "오프라인 구매 전환과 온라인 록인(Lock-in)의 연계")
cards = [
    ("즉각적 구매 동기부여",
     "전 제품 50%~최대 58% 파격 팝업 한정 할인가로 가격 장벽을 무너뜨려 현장 구매를 강력히 유도."),
    ("구매 넛지(Nudge) 설계",
     "‘1개 구매 시 티코스터, 2개 이상 구매 시 스크런치 추가’ 단계별 사은품으로 자연스러운 객단가 상승 유도."),
    ("온라인 재유입 전략",
     "럭키드로우 상품으로 ‘자사몰 1만원 할인 쿠폰’을 설계, 팝업 종료 후에도 온라인 채널로 지속 유입되는 록인 장치 마련."),
]
cw = (SW - 1.9 - 2*0.28) / 3
for i, (h, b) in enumerate(cards):
    ucard(s, 0.95 + i * (cw + 0.28), 2.55, cw, 2.0, h, b, dark=False, accent=(i == 2))
# stat tile (dark, Apple big-number style)
rrect(s, 0.95, 4.85, 7.4, 2.0, fill=C["parchment"], radius=0.06)
txt(s, 1.3, 5.1, 4.0, 0.4, [("팝업 한정 할인율", F_T, 13, C["ink48"], True, -0.24)])
txt(s, 1.28, 5.45, 6.5, 1.2,
    [("50–58", F_D, 58, C["ink"], True, -1.5), ("%", F_D, 30, C["primary"], True, -0.5)])
txt(s, 1.3, 6.45, 6.5, 0.3, [("전 제품 기준 최대 할인 · 현장 구매 전환 극대화", F_T, 12, C["ink48"], False, -0.24)])
photo(s, 8.6, 4.85, 3.78, 2.0, "샴푸·티코스터·스크런치 언박싱 컷")
footer(s, 8)

# =====================================================================  SLIDE 9  (dark tile — conclusion)
s = slide(C["tile1"])
eyebrow(s, 0.95, 0.78, "CONCLUSION", dark=True)
txt(s, 0.93, 1.18, 11.4, 0.85, [("결론 및 종합 의견", F_D, 36, C["on_dark"], True, -0.6)])
txt(s, 0.95, 1.98, 11.4, 0.5,
    [("직접 체험하며 느낀 점 및 총평", F_T, 15, C["body_muted"], False, -0.37)])
line(s, 0.95, 2.55, SW - 1.9, color=C["tile2"])
# quote card 1
rrect(s, 0.95, 2.85, 5.62, 3.5, fill=C["tile2"], radius=0.06)
txt(s, 1.25, 3.0, 1.0, 0.7, [("“", F_D, 46, C["primary_on_dark"], True, 0)])
txt(s, 1.25, 3.6, 5.05, 2.6,
    [("외부 자극·스트레스로 두피 고민이 많던 소비자 입장에서 매우 만족스러운 경험. 기능성 탈모 샴푸는 "
      "향이 딱딱하고 모발이 뻣뻣해진다는 편견이 있었으나, ‘풍성한 거품·찰랑이는 머릿결·세련된 머스크 향’으로 "
      "오프라인 체험이 준 신뢰감을 그대로 이어감.", F_T, 12.5, C["on_dark"], False, -0.24)], ls=1.42)
# quote card 2 (accent)
rrect(s, 6.73, 2.85, 5.62, 3.5, fill=C["tile2"], radius=0.06)
rrect(s, 7.03, 3.05, 0.42, 0.1, fill=C["primary_on_dark"], radius=0.5)
txt(s, 7.03, 3.25, 5.05, 2.9,
    [("화려한 마케팅 수식어 대신 제품의 본질(Pure Origin)을 강조하고, 이를 성수동이라는 감각적 오프라인 "
      "공간 안에서 영리한 혜택 체계(Real Change)로 풀어냄.  ", F_T, 12.5, C["on_dark"], False, -0.24),
     ("인디 브랜드가 고객 접점을 극대화한 훌륭한 리테일 전략 사례.",
      F_T, 12.5, C["primary_on_dark"], True, -0.24)], ls=1.42)
txt(s, 0.95, 6.7, 8.0, 0.4,
    [("Pure Origin, Real Change", F_T, 13, C["body_muted"], False, -0.24)],
    anchor=MSO_ANCHOR.MIDDLE)
footer(s, 9, dark=True)

prs.save("moet_onview_popup_report_apple.pptx")
print("saved: moet_onview_popup_report_apple.pptx  /  slides:", len(prs.slides._sldIdLst))
