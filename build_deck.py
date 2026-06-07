# -*- coding: utf-8 -*-
"""
moét X ONVIEW 팝업스토어 분석 보고서
Coinbase 디자인 시스템(다크 히어로 밴드 / Coinbase Blue 액센트 / pill 지오메트리 /
24px 라운드 카드 / 무게 400 디스플레이 / 에디토리얼 여백)을 적용한 .pptx 생성기.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.enum.text import MSO_AUTO_SIZE

# ---------------------------------------------------------------- design tokens
C = {
    "primary":        "0052FF",  # Coinbase Blue
    "primary_active": "003ECC",
    "accent_yellow":  "F4B000",
    "canvas":         "FFFFFF",
    "surface_soft":   "F7F7F7",
    "surface_strong": "EEF0F3",
    "surface_dark":   "0A0B0D",
    "surface_dark_el":"16181C",
    "hairline":       "DEE1E6",
    "ink":            "0A0B0D",
    "body":           "5B616E",
    "muted":          "7C828A",
    "muted_soft":     "A8ACB3",
    "on_primary":     "FFFFFF",
    "on_dark":        "FFFFFF",
    "on_dark_soft":   "A8ACB3",
    "up":             "05B169",
    "down":           "CF202F",
}

# Inter-substitute for Korean rendering; PowerPoint substitutes if absent.
F_DISPLAY = "Pretendard"      # CoinbaseDisplay 대체
F_SANS    = "Pretendard"      # CoinbaseSans 대체
F_MONO    = "JetBrains Mono"  # CoinbaseMono 대체

def rgb(hexs): return RGBColor.from_string(hexs)

EMU_IN = 914400
SW, SH = 13.333, 7.5  # 16:9

prs = Presentation()
prs.slide_width  = Emu(int(SW * EMU_IN))
prs.slide_height = Emu(int(SH * EMU_IN))
BLANK = prs.slide_layouts[6]


# ----------------------------------------------------------------- primitives
def slide(bg=C["canvas"]):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    r.fill.solid(); r.fill.fore_color.rgb = rgb(bg)
    r.line.fill.background()
    r.shadow.inherit = False
    return s

def _no_shadow(sp):
    sp.shadow.inherit = False

def _set_radius(sp, frac):
    """Set rounded-rect corner radius as fraction of the shorter side."""
    try:
        sp.adjustments[0] = frac
    except Exception:
        pass

def rrect(s, x, y, w, h, fill=None, line=None, line_w=1.0, radius=0.06, shadow=False):
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                            Inches(x), Inches(y), Inches(w), Inches(h))
    _set_radius(sp, radius)
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = rgb(line); sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    if shadow:
        _soft_shadow(sp)
    return sp

def rect(s, x, y, w, h, fill):
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                            Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp

def line(s, x, y, w, color=C["hairline"], weight=1.0):
    sp = s.shapes.add_connector(2, Inches(x), Inches(y), Inches(x + w), Inches(y))
    sp.line.color.rgb = rgb(color); sp.line.width = Pt(weight)
    sp.shadow.inherit = False
    return sp

def _soft_shadow(sp):
    """0 4px 12px rgba(0,0,0,0.04) — the single shadow tier."""
    spPr = sp._element.spPr
    el = spPr.makeelement(qn('a:effectLst'), {})
    sh = el.makeelement(qn('a:outerShdw'),
                        {'blurRad': '152400', 'dist': '50800',
                         'dir': '5400000', 'rotWithShape': '0'})
    clr = sh.makeelement(qn('a:srgbClr'), {'val': '000000'})
    alpha = clr.makeelement(qn('a:alpha'), {'val': '9000'})
    clr.append(alpha); sh.append(clr); el.append(sh); spPr.append(el)

def _spc(run, px):
    """letter-spacing in points (negative = tighter)."""
    run.font._rPr.set('spc', str(int(px * 100)))

def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        wrap=True, sp_after=0, line_spacing=None):
    """runs: list of (text, font, size, color, bold, tracking_pt) ; '\n' splits paragraphs."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    for m in ("left", "right", "top", "bottom"):
        setattr(tf, "margin_" + m, 0)
    first = True
    for item in runs:
        text, font, size, color, bold, tracking = item
        for li, seg in enumerate(text.split("\n")):
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = align
            if sp_after: p.space_after = Pt(sp_after)
            if line_spacing: p.line_spacing = line_spacing
            r = p.add_run(); r.text = seg
            r.font.name = font; r.font.size = Pt(size); r.font.bold = bold
            r.font.color.rgb = rgb(color)
            # apply CJK + latin font binding
            rPr = r.font._rPr
            for tag in ('a:latin', 'a:ea', 'a:cs'):
                e = rPr.find(qn(tag))
                if e is None:
                    e = rPr.makeelement(qn(tag), {}); rPr.append(e)
                e.set('typeface', font)
            if tracking:
                _spc(r, tracking)
    return tb

def badge(s, x, y, label, bg=C["surface_strong"], fg=C["ink"]):
    w = 0.13 * len(label) + 0.55
    h = 0.34
    rrect(s, x, y, w, h, fill=bg, radius=0.5)
    txt(s, x, y - 0.02, w, h, [(label, F_SANS, 10.5, fg, True, 0.5)],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return w, h

def pill_btn(s, x, y, label, bg=C["primary"], fg=C["on_primary"], w=2.0, h=0.52):
    rrect(s, x, y, w, h, fill=bg, radius=0.5)
    txt(s, x, y, w, h, [(label, F_SANS, 13, fg, True, 0)],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return w, h

def photo_ph(s, x, y, w, h, caption, dark=False):
    base = C["surface_dark_el"] if dark else C["surface_soft"]
    bord = None if dark else C["hairline"]
    fg   = C["on_dark_soft"] if dark else C["muted"]
    rrect(s, x, y, w, h, fill=base, line=bord, line_w=1.0, radius=0.05)
    # camera glyph block
    gw = 0.5
    rrect(s, x + w/2 - gw/2, y + h/2 - 0.42, gw, 0.34,
          fill=(C["surface_dark"] if dark else C["surface_strong"]), radius=0.18)
    txt(s, x + 0.25, y + h/2 + 0.04, w - 0.5, 0.7,
        [("PHOTO", F_SANS, 9, fg, True, 1.2), ("\n" + caption, F_SANS, 9.5, fg, False, 0)],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, line_spacing=1.1)

def page_num(s, n, dark=False):
    fg = C["on_dark_soft"] if dark else C["muted_soft"]
    txt(s, SW - 1.4, SH - 0.55, 1.0, 0.3,
        [(f"{n:02d} / 09", F_MONO, 9, fg, False, 0)],
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, 0.9, SH - 0.55, 4.0, 0.3,
        [("moét × ONVIEW · 성수 팝업 분석", F_SANS, 9, fg, False, 0)],
        anchor=MSO_ANCHOR.MIDDLE)

def section_header(s, num, kor, eng, subtitle, dark=False):
    ink = C["on_dark"] if dark else C["ink"]
    bodyc = C["on_dark_soft"] if dark else C["body"]
    txt(s, 0.9, 0.62, 2.0, 1.1, [(num, F_DISPLAY, 50, C["primary"], False, -2)],
        anchor=MSO_ANCHOR.TOP)
    txt(s, 2.15, 0.66, 9.5, 0.8,
        [(kor, F_DISPLAY, 30, ink, False, -1), ("   " + eng, F_SANS, 15, C["muted"], False, 0)],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, 2.17, 1.42, 10.2, 0.5, [(subtitle, F_SANS, 13.5, bodyc, False, 0)])
    line(s, 0.9, 2.0, SW - 1.8, color=(C["surface_dark_el"] if dark else C["hairline"]))

def bullet_card(s, x, y, w, h, head, body_text, accent=False):
    fill = C["canvas"]
    rrect(s, x, y, w, h, fill=fill, line=C["hairline"], line_w=1.0, radius=0.07, shadow=True)
    if accent:
        rrect(s, x, y, 0.09, h, fill=C["primary"], radius=0.0)
    pad = 0.32
    txt(s, x + pad, y + 0.26, w - 2*pad, 0.5,
        [(head, F_SANS, 14.5, C["ink"], True, 0)])
    txt(s, x + pad, y + 0.74, w - 2*pad, h - 1.0,
        [(body_text, F_SANS, 11.5, C["body"], False, 0)], line_spacing=1.3)


# =====================================================================  SLIDE 1
s = slide(C["surface_dark"])
# floating product-ui mockup stack (right)
rrect(s, 9.05, 1.55, 3.5, 4.4, fill=C["surface_dark_el"], radius=0.06)
rrect(s, 8.4, 2.55, 3.3, 3.0, fill="1E2127", radius=0.07, shadow=True)
# mock content lines inside elevated card
rrect(s, 8.75, 2.95, 1.5, 0.32, fill=C["primary"], radius=0.5)
line(s, 8.75, 3.7, 2.6, color="2A2E36", weight=2)
line(s, 8.75, 4.05, 2.6, color="2A2E36", weight=2)
line(s, 8.75, 4.4, 1.8, color="2A2E36", weight=2)
txt(s, 8.75, 4.85, 2.6, 0.5, [("PURE ORIGIN", F_MONO, 11, C["on_dark"], False, 0.5),
                              ("\nREAL CHANGE", F_MONO, 11, C["primary"], False, 0.5)],
    line_spacing=1.15)

badge(s, 0.9, 1.35, "moét × ONVIEW", bg=C["surface_dark_el"], fg=C["on_dark_soft"])
txt(s, 0.88, 2.0, 8.2, 2.6,
    [("BEAUTY\nPLAYGROUND\nIN SEONGSU", F_DISPLAY, 52, C["on_dark"], False, -1.6)],
    line_spacing=0.98)
txt(s, 0.9, 4.95, 7.6, 0.8,
    [("리테일 공간 팝업스토어 분석 보고서", F_SANS, 15, C["on_dark"], False, 0),
     ("\n모에뜨(moét) 브랜드존 중심", F_SANS, 14, C["on_dark_soft"], False, 0)],
    line_spacing=1.3)
line(s, 0.92, 6.25, 6.0, color=C["surface_dark_el"])
txt(s, 0.9, 6.45, 11.0, 0.5,
    [("14주차 과제 레포트", F_MONO, 11, C["primary"], False, 0),
     ("   ·   과목명 _______   ·   학번 _______   ·   이름 _______",
      F_SANS, 11, C["on_dark_soft"], False, 0)],
    anchor=MSO_ANCHOR.MIDDLE)

# =====================================================================  SLIDE 2
s = slide(C["canvas"])
badge(s, 0.9, 0.62, "OVERVIEW")
txt(s, 0.88, 1.05, 8.0, 0.8, [("개요", F_DISPLAY, 34, C["ink"], False, -1)])
line(s, 0.9, 1.85, SW - 1.8)
# summary (What) — left
txt(s, 0.9, 2.12, 6.3, 0.4, [("WHAT", F_SANS, 10.5, C["primary"], True, 1.0)])
txt(s, 0.9, 2.5, 6.3, 3.6,
    [("성수동 중심가에서 개최된 인디 뷰티 큐레이션 플랫폼 ‘오엔뷰(ONVIEW)’와 모에뜨의 "
      "협업 팝업스토어. K-뷰티 고관여 MZ세대를 타깃으로, 일시적 제품 판매를 넘어 "
      "오감 체험(시향·제형)과 팝업 한정 리워드를 제공하는 ‘경험형 리테일 매장’으로 "
      "기획되었음.", F_SANS, 13, C["body"], False, 0)], line_spacing=1.45)
# early-close notice
rrect(s, 0.9, 5.35, 6.3, 1.0, fill=C["surface_soft"], line=C["hairline"], radius=0.10)
rrect(s, 0.9, 5.35, 0.09, 1.0, fill=C["down"], radius=0.0)
txt(s, 1.2, 5.5, 5.9, 0.8,
    [("⚠ 운영 유의", F_SANS, 11, C["down"], True, 0),
     ("\n모에뜨 브랜드존은 6월 4일(목)까지만 조기 운영", F_SANS, 11.5, C["ink"], False, 0)],
    line_spacing=1.3)
# info table card — right
tx, ty, tw = 7.55, 2.12, 4.85
rrect(s, tx, ty, tw, 4.25, fill=C["canvas"], line=C["hairline"], radius=0.06, shadow=True)
rows = [
    ("브랜드", "모에뜨 (moét)"),
    ("팝업스토어명", "BEAUTY PLAYGROUND IN SEONGSU"),
    ("행사 기간", "2026.05.22(금) ~ 06.10(수)"),
    ("운영 시간", "11:00 ~ 19:00 (입장 마감 18:30)"),
    ("장소", "맵달SEOUL 성수 3F\n성동구 성수이로16길 5"),
]
rh = 4.25 / len(rows)
for i, (k, v) in enumerate(rows):
    ry = ty + i * rh
    if i: line(s, tx + 0.3, ry, tw - 0.6, color=C["hairline"])
    txt(s, tx + 0.32, ry, 1.5, rh, [(k, F_SANS, 10.5, C["muted"], True, 0)],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, tx + 1.75, ry, tw - 2.0, rh, [(v, F_SANS, 11.5, C["ink"], False, 0)],
        anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.15)
page_num(s, 2)

# =====================================================================  SLIDE 3
s = slide(C["surface_soft"])
section_header(s, "01", "장소성", "Locality",
               "트렌드세터의 중심지, 성수동 상권과의 유기적 결합")
cards = [
    ("트렌디한 소비층 밀집",
     "국내외 MZ세대 및 K-뷰티 고관여 소비자가 가장 활발하게 유입되는 서울 성수동 중심 상권에 위치."),
    ("우수한 접근성",
     "성수역 3번 출구에서 도보 5분 거리에 위치해 자연스러운 도보 방문객 유입률 확보."),
    ("공간의 적합성",
     "‘맵달서울’ 3층을 활용, 대형 채널에서 못 보던 신선한 인디 뷰티를 발굴·체험하는 성수동 특유의 감성과 시너지."),
]
cw = (SW - 1.8 - 2*0.3) / 3
for i, (h, b) in enumerate(cards):
    bullet_card(s, 0.9 + i * (cw + 0.3), 2.35, cw, 2.4, h, b, accent=(i == 0))
photo_ph(s, 0.9, 5.05, SW - 1.8, 1.85, "3층 행사장 진입 복도 및 입구 대기 줄")
page_num(s, 3)

# =====================================================================  SLIDE 4  (02 상징성 - 2x2)
s = slide(C["canvas"])
section_header(s, "02", "상징성", "Symbolism",
               "상품 · 컬러 · 컨셉 · 브랜드 메시지로 구축한 브랜드 아이덴티티")
quad = [
    ("[1] 상품  Products",
     "시그니처 스칼프 샴푸(톳수·창포뿌리줄기수 베이스, 식약처 허가 최대치 식물성 카페인)와 "
     "두피 앰플(마이크로 스피큘 니들샷 공법)을 중앙 매대에 배치해 고기능성 기술력을 직관적으로 체험.", False),
    ("[2] 컬러  Color",
     "메인 스카이 블루(청량함·두피 회복)로 시선을 사로잡고, 미니멀 화이트&그레이(정직함·과학적 근거)로 "
     "패키지·베이스 톤을 구성해 신뢰감 있는 무드 조성.", False),
    ("[3] 컨셉  Concept",
     "원료 설명 카드·감성 엽서로 ‘정직한 성분 연구소’를 전달. 스크런치 착용 안내와 티코스터 증정을 통해 "
     "‘산뜻한 데일리 루틴’을 경험으로 소유하게 설계.", False),
    ("[4] 브랜드 메시지  Message",
     "“Pure Origin, Real Change” — 순수한 핵심 원료(Pure Origin)를 고함량 설계해 실제 두피 개선과 "
     "긍정적 효과(Real Change)를 체감하게 한다는 약속.", True),
]
cw = (SW - 1.8 - 0.3) / 2
ch = 1.9
for i, (h, b, acc) in enumerate(quad):
    cx = 0.9 + (i % 2) * (cw + 0.3)
    cy = 2.3 + (i // 2) * (ch + 0.28)
    bullet_card(s, cx, cy, cw, ch, h, b, accent=acc)
page_num(s, 4)

# =====================================================================  SLIDE 5
s = slide(C["surface_soft"])
section_header(s, "03", "테마성", "Thematic",
               "‘Pure Origin, Real Change’ — 체험 중심의 고기능성 헤어 랩(Lab)")
cards = [
    ("메인 테마의 구체화",
     "‘뷰티 놀이터(Playground)’라는 전체 테마 속에서 ‘자극 없는 일상 속 두피 회복 솔루션’을 서브 테마로 명확히 제안."),
    ("콘셉추얼한 매대 구성",
     "단순 진열·판매를 넘어, 방문객이 자신의 두피 고민(열감·모발 빠짐·유분)을 인지하고 해답을 스스로 탐색하는 서사적 공간으로 연출."),
]
cw = (SW - 1.8 - 0.3) / 2
for i, (h, b) in enumerate(cards):
    bullet_card(s, 0.9 + i * (cw + 0.3), 2.35, cw, 2.4, h, b, accent=(i == 0))
photo_ph(s, 0.9, 5.05, SW - 1.8, 1.85, "원료 설명 리플렛·감성 엽서가 함께 디스플레이된 매대")
page_num(s, 5)

# =====================================================================  SLIDE 6
s = slide(C["canvas"])
section_header(s, "04", "콘텐츠", "Contents",
               "방문객의 오감을 자극하는 참여형 저니(Journey) 설계")
cards = [
    ("원료 스토리텔링 · 시향",
     "특허 ‘톳인삼농축액’, 고함량 식물성 카페인을 코칭하고, 세련된 플로럴 머스크 향을 직접 맡게 하는 청각·후각 콘텐츠."),
    ("촉각적 제형 테스트",
     "마이크로 스피큘(니들샷) 두피 앰플을 손등에 발라보게 해 ‘끈적임 없는 흡수력’과 ‘즉각적 쿨링감’을 현장 체감."),
    ("엔터테인먼트 요소",
     "카카오톡 채널 추가·인스타그램 팔로우 시 ‘꽝 없는 100% 당첨 럭키드로우’ 뽑기 게임으로 참여 재미 극대화."),
]
cw = (SW - 1.8 - 2*0.3) / 3
for i, (h, b) in enumerate(cards):
    bullet_card(s, 0.9 + i * (cw + 0.3), 2.35, cw, 2.4, h, b, accent=(i == 1))
photo_ph(s, 0.9, 5.05, SW - 1.8, 1.85, "두피 앰플 테스팅 장면 및 럭키드로우 돌리기 기계")
page_num(s, 6)

# =====================================================================  SLIDE 7
s = slide(C["surface_soft"])
section_header(s, "05", "화제성", "Buzz",
               "‘혜자 팝업’ 입소문과 SNS 자발적 바이럴 효과")
cards = [
    ("얼리버드 유입 (오픈런)",
     "“다양한 인디 뷰티 본품 혜택과 다채로운 이벤트”라는 정보가 온라인 뷰티 커뮤니티·SNS로 확산되며 오전 11시 오픈 전부터 대기 수요 발생."),
    ("공유 가치가 있는 굿즈",
     "소장 가치 높은 티코스터와 세련된 스크런치 증정으로, 방문객이 자발적으로 #모에뜨 #성수팝업 인증샷을 업로드하도록 유도."),
]
cw = (SW - 1.8 - 0.3) / 2
for i, (h, b) in enumerate(cards):
    bullet_card(s, 0.9 + i * (cw + 0.3), 2.35, cw, 2.4, h, b, accent=(i == 0))
photo_ph(s, 0.9, 5.05, SW - 1.8, 1.85, "인스타그램 피드 캡처 · 방문 인플루언서 인증샷 레이아웃")
page_num(s, 7)

# =====================================================================  SLIDE 8
s = slide(C["canvas"])
section_header(s, "06", "유도성", "Inducement",
               "오프라인 구매 전환과 온라인 록인(Lock-in)의 연계")
cards = [
    ("즉각적 구매 동기부여",
     "전 제품 50%~최대 58% 파격 팝업 한정 할인가로 가격 장벽을 무너뜨려 현장 구매를 강력히 유도."),
    ("구매 넛지(Nudge) 설계",
     "‘1개 구매 시 티코스터, 2개 이상 구매 시 스크런치 추가’ 단계별 사은품으로 자연스러운 객단가 상승 유도."),
    ("온라인 재유입 전략",
     "럭키드로우 상품으로 ‘자사몰 1만원 할인 쿠폰’을 설계, 팝업 종료 후에도 온라인 채널로 지속 유입되는 록인 장치 마련."),
]
cw = (SW - 1.8 - 2*0.3) / 3
for i, (h, b) in enumerate(cards):
    bullet_card(s, 0.9 + i * (cw + 0.3), 2.35, cw, 2.4, h, b, accent=(i == 2))
# discount stat highlight (mono number)
photo_ph(s, 0.9, 5.05, 7.4, 1.85, "최종 구매 후 수령한 샴푸·티코스터·스크런치 언박싱 컷")
rrect(s, 8.5, 5.05, 3.9, 1.85, fill=C["surface_dark"], radius=0.07, shadow=True)
txt(s, 8.7, 5.25, 3.5, 0.4, [("팝업 한정 할인율", F_SANS, 10.5, C["on_dark_soft"], True, 0.5)])
txt(s, 8.7, 5.55, 3.5, 1.0, [("50–58", F_MONO, 40, C["on_dark"], False, 0),
                             ("%", F_MONO, 22, C["primary"], False, 0)])
txt(s, 8.7, 6.5, 3.5, 0.3, [("전 제품 기준 최대 할인", F_SANS, 10, C["on_dark_soft"], False, 0)])
page_num(s, 8)

# =====================================================================  SLIDE 9  (Conclusion - dark CTA band)
s = slide(C["surface_dark"])
badge(s, 0.9, 0.75, "CONCLUSION", bg=C["surface_dark_el"], fg=C["on_dark_soft"])
txt(s, 0.88, 1.2, 11.5, 0.9,
    [("결론 및 종합 의견", F_DISPLAY, 36, C["on_dark"], False, -1)])
txt(s, 0.9, 2.05, 11.5, 0.4,
    [("직접 체험하며 느낀 점 및 총평", F_SANS, 13.5, C["on_dark_soft"], False, 0)])
line(s, 0.92, 2.6, SW - 1.84, color=C["surface_dark_el"])

# quote card 1
rrect(s, 0.9, 2.85, 5.65, 3.55, fill=C["surface_dark_el"], radius=0.06)
txt(s, 1.2, 3.05, 0.8, 0.6, [("“", F_DISPLAY, 44, C["primary"], False, 0)])
txt(s, 1.2, 3.65, 5.1, 2.6,
    [("외부 자극·스트레스로 두피 고민이 많던 소비자 입장에서 매우 만족스러운 경험. "
      "기능성 탈모 샴푸는 향이 딱딱하고 모발이 뻣뻣해진다는 편견이 있었으나, "
      "‘풍성한 거품·찰랑이는 머릿결·세련된 머스크 향’으로 오프라인 체험이 준 신뢰감을 그대로 이어감.",
      F_SANS, 12.5, C["on_dark"], False, 0)], line_spacing=1.4)

# quote card 2 (accent)
rrect(s, 6.75, 2.85, 5.65, 3.55, fill=C["surface_dark_el"], radius=0.06)
rrect(s, 6.75, 2.85, 0.09, 3.55, fill=C["primary"], radius=0.0)
txt(s, 7.05, 3.05, 0.8, 0.6, [("“", F_DISPLAY, 44, C["primary"], False, 0)])
txt(s, 7.05, 3.65, 5.1, 2.6,
    [("화려한 마케팅 수식어 대신 제품의 본질(Pure Origin)을 강조하고, 이를 성수동이라는 "
      "감각적 오프라인 공간 안에서 영리한 혜택 체계(Real Change)로 풀어냄. ",
      F_SANS, 12.5, C["on_dark"], False, 0),
     ("인디 브랜드가 고객 접점을 극대화한 훌륭한 리테일 전략 사례.",
      F_SANS, 12.5, C["primary"], True, 0)], line_spacing=1.4)
txt(s, 0.9, 6.75, 8.0, 0.4,
    [("Pure Origin, Real Change", F_MONO, 12, C["on_dark_soft"], False, 0.5)],
    anchor=MSO_ANCHOR.MIDDLE)
page_num(s, 9, dark=True)

prs.save("moet_onview_popup_report.pptx")
print("saved: moet_onview_popup_report.pptx  /  slides:", len(prs.slides._sldIdLst))
