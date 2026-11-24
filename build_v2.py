#!/usr/bin/env python3
"""gosengmin_1124_2.pptx 빌드 — 대본 v2 기준 재구성
S01~S08, S12~S14: 기존 유지 | S09: 매출+RLUSD 투업 | S10: 슈워츠 | S11: 프랭클린(신규)"""
import shutil
from pptx import Presentation
from pptx.util import Inches

shutil.copy('gosengmin_1124_1.pptx', 'gosengmin_1124_2.pptx')
p = Presentation('gosengmin_1124_2.pptx')
slides = list(p.slides)

def pics(s): return [sh for sh in s.shapes if sh.shape_type == 13]

# ── S09 (idx 8): s5_money 이동 + s10_rlusd 추가 (투업, 페이지번호 12.05 회피)
s = slides[8]
ps = pics(s); assert len(ps) == 1, "S09 구조 이상"
ps[0].left, ps[0].top = Inches(0.90), Inches(0.47)
ps[0].width, ps[0].height = Inches(5.35), Inches(6.56)
s.shapes.add_picture('crop/s10_rlusd.png', Inches(6.65), Inches(0.47), Inches(5.35), Inches(6.56))

# ── S10 (idx 9): rlusd → schwartz
s = slides[9]
ps = pics(s)
ps[0]._element.getparent().remove(ps[0]._element)
s.shapes.add_picture('crop/s7_schwartz.png', Inches(3.77), Inches(0.20), Inches(5.79), Inches(7.10))

# ── S11 (idx 10): schwartz → franklin (aspect 1.02, 높이 7.10 고정 → 폭 7.24)
s = slides[10]
ps = pics(s)
ps[0]._element.getparent().remove(ps[0]._element)
s.shapes.add_picture('crop/s11_franklin.png', Inches(3.05), Inches(0.20), Inches(7.24), Inches(7.10))

p.save('gosengmin_1124_2.pptx')
print('빌드 완료: gosengmin_1124_2.pptx')

# ── 검증: 슬라이드별 이미지 재추출
import hashlib, glob, os
hmap = {hashlib.md5(open(f,'rb').read()).hexdigest(): os.path.basename(f) for f in glob.glob('crop/*.png')}
p2 = Presentation('gosengmin_1124_2.pptx')
for i, sl in enumerate(p2.slides, 1):
    for sh in sl.shapes:
        if sh.shape_type == 13:
            h = hashlib.md5(sh.image.blob).hexdigest()
            n = hmap.get(h, f"?{h[:6]}")
            print(f"S{i:02d} <- {n:22s} pos=({sh.left/914400:.2f},{sh.top/914400:.2f}) {sh.width/914400:.2f}x{sh.height/914400:.2f}in")
