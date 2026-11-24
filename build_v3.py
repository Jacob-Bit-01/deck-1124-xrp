#!/usr/bin/env python3
"""gosengmin_1124_3.pptx — 핵심 크롭 v3 (영상 가독성 개선)
순차 슬라이드 추가(1~14) — 순서 버그 수정본. 배치: w=min(12.2, aspect*6.4)"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from PIL import Image
import os

IMG = {2:'crop2/c_fgi.png', 3:'crop2/s3_worstnov.png', 4:'crop2/s1_jiseon.png',
       5:'crop2/s2_tomlee.png', 6:'crop2/s8_escrow.png', 7:'crop2/c_xrpscan.png',
       8:'crop2/s4_ssook.png', 10:'crop2/s7_schwartz.png', 11:'crop2/s11_franklin.png',
       12:'crop2/s9_whales.png', 13:'crop2/s11_yna.png'}
S9_PAIR = ['crop2/s5_money_h.png', 'crop2/s10_rlusd_h.png']

prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

def new_slide(n):
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tb = s.shapes.add_textbox(Inches(12.05), Inches(7.05), Inches(1.10), Inches(0.35))
    r = tb.text_frame.paragraphs[0].add_run(); r.text = f'{n} / 14'
    r.font.size = Pt(10); r.font.name = 'Apple SD Gothic Neo'; r.font.color.rgb = RGBColor(0x99,0x99,0x99)
    return s

def add_img(s, path, w_force=None):
    im = Image.open(path); a = im.size[0]/im.size[1]
    w = w_force or min(12.2, a*6.4); h = w/a
    s.shapes.add_picture(path, Inches((13.333-w)/2), Inches((7.5-h)/2), Inches(w), Inches(h))

def big_text(s, lines, box=(1.2, 0.8, 11.0, 5.9)):
    tb = s.shapes.add_textbox(Inches(box[0]), Inches(box[1]), Inches(box[2]), Inches(box[3]))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, (txt, size, bold, col) in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.space_after = Pt(16)
        run = para.add_run(); run.text = txt
        run.font.size = Pt(size); run.font.bold = bold; run.font.name = 'Apple SD Gothic Neo'
        run.font.color.rgb = RGBColor(col, col, col)

# ── 순차 빌드: 1~14 순서대로
for n in range(1, 15):
    s = new_slide(n)
    if n == 1:
        big_text(s, [
            ('XRP  $0.94  (₩1,267)', 48, True, 0x11),
            ('24시간  +0.48%', 32, False, 0x11),
            ('BTC $60,411 · ETH $1,696', 26, False, 0x11),
            ('공포탐욕지수 11 — 극도공포', 32, True, 0x11),
            ('2026.11.24 16:40 기준', 20, False, 0x66)])
    elif n == 9:
        W9 = 10.5
        dims = [W9/(Image.open(p).size[0]/Image.open(p).size[1]) for p in S9_PAIR]
        total = sum(dims) + 0.35; y = (7.5-total)/2
        for path, h in zip(S9_PAIR, dims):
            s.shapes.add_picture(path, Inches((13.333-W9)/2), Inches(y), Inches(W9), Inches(h))
            y += h + 0.35
    elif n == 14:
        big_text(s, [
            ('① 0.9달러 4차 방어 여부', 30, True, 0x11),
            ('② RLUSD 30억 달러 돌파 — XRP 수요 전이', 30, True, 0x11),
            ('③ 큰손 보유 68.5% — 유동 공급 압박', 30, True, 0x11)], box=(1.2, 1.5, 11.0, 4.5))
    else:
        add_img(s, IMG[n], w_force=(9.5 if n == 8 else None))

prs.save('gosengmin_1124_3.pptx')
print('빌드 완료:', round(os.path.getsize('gosengmin_1124_3.pptx')/1e6, 2), 'MB')

# 구조 검증 — 위치 순서대로
from pptx import Presentation as P2
p3 = P2('gosengmin_1124_3.pptx')
import hashlib, glob
hmap = {hashlib.md5(open(f,'rb').read()).hexdigest(): os.path.basename(f) for f in glob.glob('crop2/*.png')}
for i, sl in enumerate(p3.slides, 1):
    pss = [sh for sh in sl.shapes if sh.shape_type == 13]
    names = [hmap.get(hashlib.md5(sh.image.blob).hexdigest(), '?') for sh in pss]
    print(f"S{i:02d} {'+'.join(names) if names else 'TEXT'}")
