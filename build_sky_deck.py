# -*- coding: utf-8 -*-
"""
moét X ONVIEW 팝업스토어 분석 보고서 — Sky Blue 템플릿 + 실사 이미지 배치본.
Apple식 풀블리드 타일 구조를 유지하되 액센트/밴드를 모에뜨 스카이블루로 리컬러.
"""
import glob, hashlib, os
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- sky-blue tokens
C = {
    "sky":        "1E91CF",   # primary accent (CTA/eyebrow/accent bar)
    "sky_deep":   "0E6CA6",   # strong accent text on light
    "sky_link":   "7FC9EE",   # accent on dark surfaces
    "tint":       "EAF5FB",   # light band surface
    "tint2":      "DCEEF8",   # secondary tint
    "pale":       "F4FAFD",
    "canvas":     "FFFFFF",
    "navy":       "0E2E45",   # dark closing tile
    "navy2":      "163A56",   # card on navy
    "ink":        "1B2A36",   # cool near-black text
    "ink60":      "55636E",   # body
    "ink40":      "8A949C",   # muted / fine print
    "on_dark":    "FFFFFF",
    "muted_dark": "B9CBD8",
    "hairline":   "D7E3EC",
}
F_D = "Pretendard"
F_T = "Pretendard"

def rgb(h): return RGBColor.from_string(h)
EMU_IN = 914400
SW, SH = 13.333, 7.5

SRC = sorted(glob.glob('assets_raw/*.jpg'))
os.makedirs('assets_proc', exist_ok=True)

prs = Presentation()
prs.slide_width  = Emu(int(SW * EMU_IN))
prs.slide_height = Emu(int(SH * EMU_IN))
BLANK = prs.slide_layouts[6]

# ----------------------------------------------------------------- primitives
def slide(bg):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    r.fill.solid(); r.fill.fore_color.rgb = rgb(bg); r.line.fill.background()
    r.shadow.inherit = False
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

def line(s, x, y, w, color=C["hairline"], weight=1.0):
    sp = s.shapes.add_connector(2, Inches(x), Inches(y), Inches(x + w), Inches(y))
    sp.line.color.rgb = rgb(color); sp.line.width = Pt(weight); sp.shadow.inherit = False
    return sp

def _shadow(el):
    """Soft product/photo shadow."""
    spPr = el.spPr
    e = spPr.makeelement(qn('a:effectLst'), {})
    sh = e.makeelement(qn('a:outerShdw'),
                       {'blurRad': '180000', 'dist': '45000', 'dir': '5400000', 'rotWithShape': '0'})
    clr = sh.makeelement(qn('a:srgbClr'), {'val': '13344A'})
    clr.append(clr.makeelement(qn('a:alpha'), {'val': '26000'}))
    sh.append(clr); e.append(sh); spPr.append(e)

def _round_pic(pic, frac):
    spPr = pic._element.spPr
    for tag in ('a:prstGeom', 'a:custGeom'):
        ex = spPr.find(qn(tag))
        if ex is not None: spPr.remove(ex)
    geom = spPr.makeelement(qn('a:prstGeom'), {'prst': 'roundRect'})
    av = geom.makeelement(qn('a:avLst'), {})
    av.append(av.makeelement(qn('a:gd'), {'name': 'adj', 'fmla': f'val {int(frac*100000)}'}))
    geom.append(av); spPr.append(geom)

def _prep(idx, tw, th, mode):
    """Return (path, w_in, h_in) processed for the box."""
    src = SRC[idx]
    im = ImageOps.exif_transpose(Image.open(src)).convert('RGB')
    iw, ih = im.size
    if mode == 'cover':
        tr = tw / th; ir = iw / ih
        if ir > tr:
            nw = int(ih * tr); im = im.crop(((iw - nw)//2, 0, (iw - nw)//2 + nw, ih))
        else:
            nh = int(iw / tr); im = im.crop((0, (ih - nh)//2, iw, (ih - nh)//2 + nh))
        out_w, out_h = tw, th
    else:  # contain
        scale = min(tw/iw, th/ih)
        out_w, out_h = iw*scale, ih*scale
    key = hashlib.md5(f'{src}{tw:.2f}{th:.2f}{mode}'.encode()).hexdigest()[:10]
    path = f'assets_proc/{key}.jpg'
    im.save(path, quality=88)
    return path, out_w, out_h

def add_image(s, idx, x, y, w, h, mode='cover', radius=0.05, shadow=True, caption=None):
    path, pw, ph = _prep(idx, w, h, mode)
    px = x + (w - pw)/2; py = y + (h - ph)/2
    pic = s.shapes.add_picture(path, Inches(px), Inches(py), Inches(pw), Inches(ph))
    _round_pic(pic, radius)
    if shadow: _shadow(pic._element)
    if caption:
        txt(s, x, y + h + 0.06, w, 0.3, [(caption, F_T, 9, C["ink40"], False, -0.2)],
            align=PP_ALIGN.CENTER)
    return pic

def _spc(run, px): run.font._rPr.set('spc', str(int(px * 100)))

def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, wrap=True, ls=None):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE; tf.vertical_anchor = anchor
    for m in ("left", "right", "top", "bottom"): setattr(tf, "margin_" + m, 0)
    first = True
    for text, font, size, color, bold, tracking in runs:
        for seg in text.split("\n"):
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False; p.alignment = align
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

def pill(s, x, y, label, bg=C["sky"], fg=C["on_dark"], w=1.9, h=0.5, outline=None):
    rrect(s, x, y, w, h, fill=bg, lc=outline, lw=1.2, radius=0.5)
    txt(s, x, y, w, h, [(label, F_T, 13, fg, True, -0.2)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def eyebrow(s, x, y, text, dark=False):
    col = C["sky_link"] if dark else C["sky_deep"]
    txt(s, x, y, 9.0, 0.34, [(text, F_T, 13, col, True, -0.1)])

def footer(s, n, dark=False):
    fg = C["muted_dark"] if dark else C["ink40"]
    txt(s, 0.95, SH - 0.52, 8.0, 0.3,
        [("moét × ONVIEW · BEAUTY PLAYGROUND IN SEONGSU", F_T, 9.5, fg, False, -0.12)],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, SW - 2.0, SH - 0.52, 1.05, 0.3, [(f"{n:02d} / 09", F_T, 9.5, fg, False, -0.12)],
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

def header(s, num, kor, eng, subtitle, dark=False):
    ink = C["on_dark"] if dark else C["ink"]
    sub = C["muted_dark"] if dark else C["ink60"]
    eyebrow(s, 0.95, 0.7, f"{num}  {eng.upper()}", dark=dark)
    txt(s, 0.93, 1.1, 11.4, 0.8, [(kor, F_D, 35, ink, True, -0.6)])
    txt(s, 0.95, 1.88, 11.4, 0.5, [(subtitle, F_T, 15, sub, False, -0.37)], ls=1.2)
    line(s, 0.95, 2.42, SW - 1.9, color=(C["navy2"] if dark else C["hairline"]))

def ucard(s, x, y, w, h, head, body, dark=False, accent=False):
    if dark:
        rrect(s, x, y, w, h, fill=C["navy2"], radius=0.07); hc, bc = C["on_dark"], C["muted_dark"]
        acol = C["sky_link"]
    else:
        rrect(s, x, y, w, h, fill=C["canvas"], lc=C["hairline"], lw=1.0, radius=0.07)
        hc, bc = C["ink"], C["ink60"]; acol = C["sky"]
    pad = 0.32
    if accent:
        rrect(s, x + pad, y + 0.3, 0.42, 0.1, fill=acol, radius=0.5); ho = 0.5
    else: ho = 0.3
    txt(s, x + pad, y + ho, w - 2*pad, 0.5, [(head, F_T, 14.5, hc, True, -0.37)])
    txt(s, x + pad, y + ho + 0.46, w - 2*pad, h - ho - 0.66,
        [(body, F_T, 11.5, bc, False, -0.24)], ls=1.3)

def hcard(s, x, y, w, h, idx, head, body, accent=False):
    rrect(s, x, y, w, h, fill=C["canvas"], lc=C["hairline"], lw=1.0, radius=0.07)
    pad = 0.28
    isz = h - 2*pad
    add_image(s, idx, x + pad, y + pad, isz, isz, mode='cover', radius=0.06, shadow=False)
    tx = x + pad + isz + 0.28
    tw = w - (tx - x) - pad
    if accent:
        rrect(s, tx, y + 0.34, 0.4, 0.09, fill=C["sky"], radius=0.5); ho = 0.52
    else: ho = 0.34
    txt(s, tx, y + ho, tw, 0.45, [(head, F_T, 14, C["ink"], True, -0.37)])
    txt(s, tx, y + ho + 0.42, tw, h - ho - 0.6, [(body, F_T, 11, C["ink60"], False, -0.24)], ls=1.28)


# =====================================================================  SLIDE 1
s = slide(C["tint"])
# soft sky panel left accent
txt(s, 0.5, 1.5, 7.7, 0.4, [("팝업 리테일 공간 분석 보고서", F_T, 15, C["sky_deep"], True, -0.1)],
    align=PP_ALIGN.LEFT)
txt(s, 0.92, 2.0, 7.5, 2.2, [("BEAUTY\nPLAYGROUND\nIN SEONGSU", F_D, 52, C["ink"], True, -1.4)], ls=1.02)
txt(s, 0.95, 4.55, 7.2, 0.9,
    [("리테일 공간 팝업스토어 분석 보고서", F_T, 18, C["ink60"], False, -0.2),
     ("\n모에뜨(moét) 브랜드존 중심 · moét × ONVIEW", F_T, 14, C["ink40"], False, -0.2)], ls=1.35)
pill(s, 0.95, 5.75, "분석 개요", bg=C["sky"], fg=C["on_dark"], w=1.8, h=0.5)
pill(s, 2.9, 5.75, "종합 의견", bg=None, fg=C["sky_deep"], w=1.8, h=0.5, outline=C["sky"])
txt(s, 0.95, 6.65, 7.2, 0.4,
    [("14주차 과제 레포트 · 과목명 _____ · 학번 _____ · 이름 _____", F_T, 11.5, C["ink40"], False, -0.12)])
# official key visual (contain, no crop)
add_image(s, 22, 8.55, 0.85, 4.1, 5.8, mode='contain', radius=0.03, shadow=True)
footer(s, 1)

# =====================================================================  SLIDE 2
s = slide(C["canvas"])
eyebrow(s, 0.95, 0.7, "OVERVIEW")
txt(s, 0.93, 1.1, 11.4, 0.8, [("개요", F_D, 35, C["ink"], True, -0.6)])
line(s, 0.95, 1.98, SW - 1.9)
txt(s, 0.95, 2.24, 6.1, 0.4, [("WHAT", F_T, 12, C["sky_deep"], True, -0.12)])
txt(s, 0.95, 2.62, 6.1, 3.0,
    [("성수동 중심가에서 개최된 인디 뷰티 큐레이션 플랫폼 ‘오엔뷰(ONVIEW)’와 모에뜨의 협업 "
      "팝업스토어. K-뷰티 고관여 MZ세대를 타깃으로, 일시적 제품 판매를 넘어 오감 체험(시향·제형)과 "
      "팝업 한정 리워드를 제공하는 ‘경험형 리테일 매장’으로 기획되었음.", F_T, 15.5, C["ink60"], False, -0.37)], ls=1.5)
rrect(s, 0.95, 5.5, 6.1, 0.95, fill=C["tint"], radius=0.10)
txt(s, 1.25, 5.66, 5.6, 0.7,
    [("운영 유의", F_T, 12, C["sky_deep"], True, -0.2),
     ("\n모에뜨 브랜드존은 6월 4일(목)까지만 조기 운영", F_T, 13, C["ink"], False, -0.37)], ls=1.3)
# right column: photo + info table
add_image(s, 0, 7.5, 2.24, 4.88, 1.95, mode='cover', radius=0.05, shadow=True)
tx, ty, tw = 7.5, 4.42, 4.88
rrect(s, tx, ty, tw, 2.0, fill=C["pale"], lc=C["hairline"], lw=1.0, radius=0.07)
rows = [
    ("브랜드", "모에뜨 (moét)"),
    ("기간", "2026.05.22(금) ~ 06.10(수)"),
    ("운영 시간", "11:00 ~ 19:00 (입장 마감 18:30)"),
    ("장소", "맵달SEOUL 성수 3F · 성수이로16길 5"),
]
rh = 2.0 / len(rows)
for i, (k, v) in enumerate(rows):
    ry = ty + i * rh
    if i: line(s, tx + 0.3, ry, tw - 0.6, color=C["hairline"])
    txt(s, tx + 0.32, ry, 1.35, rh, [(k, F_T, 10.5, C["ink40"], False, -0.24)], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, tx + 1.7, ry, tw - 1.95, rh, [(v, F_T, 11, C["ink"], True, -0.24)], anchor=MSO_ANCHOR.MIDDLE, ls=1.1)
footer(s, 2)

# =====================================================================  SLIDE 3  (Locality)
s = slide(C["tint"])
header(s, "01", "장소성", "Locality", "트렌드세터의 중심지, 성수동 상권과의 유기적 결합")
cards = [
    ("트렌디한 소비층 밀집", "국내외 MZ세대 및 K-뷰티 고관여 소비자가 가장 활발히 유입되는 서울 성수동 중심 상권에 위치."),
    ("우수한 접근성", "성수역 3번 출구에서 도보 5분 거리에 위치해 자연스러운 도보 방문객 유입률 확보."),
    ("공간의 적합성", "‘맵달서울’ 3층을 활용, 대형 채널에서 못 보던 신선한 인디 뷰티를 발굴·체험하는 성수동 감성과 시너지."),
]
cw = (SW - 1.9 - 2*0.28) / 3
for i, (h, b) in enumerate(cards):
    ucard(s, 0.95 + i * (cw + 0.28), 2.62, cw, 1.95, h, b, accent=(i == 0))
iw = (SW - 1.9 - 0.3) / 2
add_image(s, 24, 0.95, 4.78, iw, 1.78, mode='cover', caption="맵달SEOUL 성수 외관")
add_image(s, 23, 0.95 + iw + 0.3, 4.78, iw, 1.78, mode='contain', caption="Location Map · 성수")
footer(s, 3)

# =====================================================================  SLIDE 4  (Symbolism)
s = slide(C["canvas"])
header(s, "02", "상징성", "Symbolism", "상품 · 컬러 · 컨셉 · 브랜드 메시지로 구축한 브랜드 아이덴티티")
quad = [
    (20, "[1] 상품  Products",
     "스칼프 샴푸(톳수·창포뿌리줄기수 베이스, 고함량 식물성 카페인)와 두피 앰플(마이크로 스피큘 공법)을 "
     "중앙 매대에 배치, 고기능성 기술력을 직관적으로 체험.", False),
    (7, "[2] 컬러  Color",
     "메인 스카이 블루(청량함·두피 회복)로 시선을 사로잡고, 미니멀 화이트&그레이(정직함·과학적 근거)로 "
     "신뢰감 있는 무드 조성.", False),
    (14, "[3] 컨셉  Concept",
     "원료 설명 카드·감성 엽서로 ‘정직한 성분 연구소’를 전달. 스크런치·티코스터 증정으로 ‘산뜻한 데일리 "
     "루틴’을 경험으로 소유하게 설계.", False),
    (16, "[4] 메시지  Message",
     "“Pure Origin, Real Change” — 순수한 핵심 원료를 고함량 설계해 실제 두피 개선과 긍정적 효과를 "
     "체감하게 한다는 약속.", True),
]
cw = (SW - 1.9 - 0.28) / 2
ch = 1.96
for i, (idx, h, b, acc) in enumerate(quad):
    cx = 0.95 + (i % 2) * (cw + 0.28)
    cy = 2.6 + (i // 2) * (ch + 0.22)
    hcard(s, cx, cy, cw, ch, idx, h, b, accent=acc)
footer(s, 4)

# =====================================================================  SLIDE 5  (Thematic)
s = slide(C["tint"])
header(s, "03", "테마성", "Thematic", "‘Pure Origin, Real Change’ — 체험 중심의 고기능성 헤어 랩(Lab)")
cards = [
    ("메인 테마의 구체화", "‘뷰티 놀이터(Playground)’ 속에서 ‘자극 없는 일상 속 두피 회복 솔루션’을 서브 테마로 명확히 제안."),
    ("콘셉추얼한 매대 구성", "단순 진열·판매를 넘어, 방문객이 자신의 두피 고민을 인지하고 해답을 스스로 탐색하는 서사적 공간으로 연출."),
]
cw = (SW - 1.9 - 0.28) / 2
for i, (h, b) in enumerate(cards):
    ucard(s, 0.95 + i * (cw + 0.28), 2.62, cw, 1.95, h, b, accent=(i == 0))
iw = (SW - 1.9 - 0.3) / 2
add_image(s, 12, 0.95, 4.78, iw, 1.78, mode='cover', caption="원료 설명 리플렛")
add_image(s, 13, 0.95 + iw + 0.3, 4.78, iw, 1.78, mode='cover', caption="성분 안내 · 샘플 카드")
footer(s, 5)

# =====================================================================  SLIDE 6  (Contents)
s = slide(C["canvas"])
header(s, "04", "콘텐츠", "Contents", "방문객의 오감을 자극하는 참여형 저니(Journey) 설계")
cards = [
    ("원료 스토리텔링 · 시향", "특허 ‘톳인삼농축액’, 고함량 식물성 카페인을 코칭하고, 세련된 플로럴 머스크 향을 직접 맡게 하는 콘텐츠."),
    ("촉각적 제형 테스트", "마이크로 스피큘 두피 앰플을 손등에 발라보게 해 ‘끈적임 없는 흡수력’과 ‘즉각적 쿨링감’을 현장 체감."),
    ("엔터테인먼트 요소", "카톡 채널 추가·인스타 팔로우 시 ‘꽝 없는 100% 당첨 럭키드로우’ 게임으로 참여 재미 극대화."),
]
cw = (SW - 1.9 - 2*0.28) / 3
for i, (h, b) in enumerate(cards):
    ucard(s, 0.95 + i * (cw + 0.28), 2.62, cw, 1.95, h, b, accent=(i == 1))
iw3 = (SW - 1.9 - 2*0.3) / 3
for j, idx in enumerate([11, 6, 10]):
    add_image(s, idx, 0.95 + j * (iw3 + 0.3), 4.78, iw3, 1.78, mode='cover')
footer(s, 6)

# =====================================================================  SLIDE 7  (Buzz)
s = slide(C["tint"])
header(s, "05", "화제성", "Buzz", "‘혜자 팝업’ 입소문과 SNS 자발적 바이럴 효과")
cards = [
    ("얼리버드 유입 (오픈런)", "“다양한 인디 뷰티 본품 혜택·이벤트” 정보가 뷰티 커뮤니티·SNS로 확산되며 오전 11시 오픈 전부터 대기 수요 발생."),
    ("공유 가치가 있는 굿즈", "소장 가치 높은 티코스터·스크런치 증정으로, 방문객이 자발적으로 #모에뜨 #성수팝업 인증샷을 업로드하도록 유도."),
]
cw = (SW - 1.9 - 0.28) / 2
for i, (h, b) in enumerate(cards):
    ucard(s, 0.95 + i * (cw + 0.28), 2.62, cw, 1.95, h, b, accent=(i == 0))
iw3 = (SW - 1.9 - 2*0.3) / 3
for j, idx in enumerate([17, 15, 9]):
    add_image(s, idx, 0.95 + j * (iw3 + 0.3), 4.78, iw3, 1.78, mode='cover')
footer(s, 7)

# =====================================================================  SLIDE 8  (Inducement)
s = slide(C["canvas"])
header(s, "06", "유도성", "Inducement", "오프라인 구매 전환과 온라인 록인(Lock-in)의 연계")
cards = [
    ("즉각적 구매 동기부여", "전 제품 50%~최대 58% 파격 팝업 한정 할인가로 가격 장벽을 낮춰 현장 구매를 강력히 유도."),
    ("구매 넛지(Nudge)", "‘1개 구매 시 티코스터, 2개 이상 구매 시 스크런치 추가’ 단계별 사은품으로 객단가 상승 유도."),
    ("온라인 재유입", "럭키드로우 상품으로 ‘자사몰 1만원 할인 쿠폰’을 설계, 종료 후에도 온라인 채널로 지속 유입되는 록인 장치."),
]
cw = (SW - 1.9 - 2*0.28) / 3
for i, (h, b) in enumerate(cards):
    ucard(s, 0.95 + i * (cw + 0.28), 2.62, cw, 1.9, h, b, accent=(i == 2))
# stat tile (sky)
rrect(s, 0.95, 4.78, 5.5, 1.78, fill=C["tint"], radius=0.07)
txt(s, 1.28, 5.0, 4.0, 0.4, [("팝업 한정 할인율", F_T, 13, C["sky_deep"], True, -0.24)])
txt(s, 1.26, 5.34, 5.0, 1.0, [("50–58", F_D, 50, C["ink"], True, -1.4), ("%", F_D, 26, C["sky"], True, -0.5)])
txt(s, 1.28, 6.22, 5.0, 0.3, [("전 제품 기준 최대 할인 · 현장 전환 극대화", F_T, 11, C["ink40"], False, -0.24)])
add_image(s, 26, 6.7, 4.78, 2.7, 1.78, mode='cover')
add_image(s, 25, 9.6, 4.78, 2.78, 1.78, mode='cover')
footer(s, 8)

# =====================================================================  SLIDE 9  (Conclusion)
s = slide(C["navy"])
eyebrow(s, 0.95, 0.76, "CONCLUSION", dark=True)
txt(s, 0.93, 1.16, 11.4, 0.85, [("결론 및 종합 의견", F_D, 35, C["on_dark"], True, -0.6)])
txt(s, 0.95, 1.96, 11.4, 0.5, [("직접 체험하며 느낀 점 및 총평", F_T, 15, C["muted_dark"], False, -0.37)])
line(s, 0.95, 2.5, SW - 1.9, color=C["navy2"])
rrect(s, 0.95, 2.8, 5.62, 3.5, fill=C["navy2"], radius=0.06)
txt(s, 1.25, 2.95, 1.0, 0.7, [("“", F_D, 46, C["sky_link"], True, 0)])
txt(s, 1.25, 3.55, 5.05, 2.6,
    [("외부 자극·스트레스로 두피 고민이 많던 소비자 입장에서 매우 만족스러운 경험. 기능성 탈모 샴푸는 "
      "향이 딱딱하고 모발이 뻣뻣해진다는 편견이 있었으나, ‘풍성한 거품·찰랑이는 머릿결·세련된 머스크 향’으로 "
      "오프라인 체험이 준 신뢰감을 그대로 이어감.", F_T, 12.5, C["on_dark"], False, -0.24)], ls=1.42)
rrect(s, 6.73, 2.8, 5.62, 3.5, fill=C["navy2"], radius=0.06)
rrect(s, 7.03, 3.0, 0.42, 0.1, fill=C["sky_link"], radius=0.5)
txt(s, 7.03, 3.2, 5.05, 2.9,
    [("화려한 마케팅 수식어 대신 제품의 본질(Pure Origin)을 강조하고, 이를 성수동이라는 감각적 오프라인 "
      "공간 안에서 영리한 혜택 체계(Real Change)로 풀어냄.  ", F_T, 12.5, C["on_dark"], False, -0.24),
     ("인디 브랜드가 고객 접점을 극대화한 훌륭한 리테일 전략 사례.", F_T, 12.5, C["sky_link"], True, -0.24)], ls=1.42)
txt(s, 0.95, 6.65, 8.0, 0.4, [("Pure Origin, Real Change", F_T, 13, C["muted_dark"], False, -0.24)],
    anchor=MSO_ANCHOR.MIDDLE)
footer(s, 9, dark=True)

prs.save("moet_onview_popup_report_sky.pptx")
print("saved: moet_onview_popup_report_sky.pptx  /  slides:", len(prs.slides._sldIdLst))
