#!/usr/bin/env python3
"""Appendix '왜 온디바이스여야 하나요?' 장(apx_new.s_ondevice)을 그려, 지금 덱의 사생활 보호 Q&A 장 바로 뒤에 끼운다.

  python tools/ondevice_apx.py                → out/ondevice/one.pptx(그 장 하나), Weave_통합본_Claude_온디바이스.pptx(지금 덱 + 그 장)
  python tools/ondevice_apx.py --inplace      → 지금 덱(Weave_통합본_Claude.pptx)에 바로 끼움(PowerPoint에서 닫아 둔 뒤에)
  python tools/ondevice_apx.py --png          → 끼운 덱의 그 장을 PowerPoint로 찍음(out/ondevice/slide.png, Windows)

make.py를 다시 돌리면(원본/의 덱들이 필요) appendix_specs가 이 장을 같은 자리에 그리므로 이 도구는 그때까지 쓰는 것.
Windows에서도 돌게: 영상 렌더에만 쓰는 fcntl은 빈 모듈로, 글자 폭을 재는 SamsungOneKorean은 06_에셋/글꼴에 없으면 설치된 TTF로.
"""
import argparse
import copy
import subprocess
import sys
import types
from pathlib import Path

if sys.platform == 'win32':
    sys.modules.setdefault('fcntl', types.ModuleType('fcntl'))

HERE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / 'tools'))
sys.path.insert(0, str(HERE.parent / '09_HowItWorks' / '장표5_claude'))

import hiw_lib as L  # noqa: E402

if not (L.FONTDIR / L.SAMSUNG_FILE[700]).exists():
    for d in (Path('C:/Windows/Fonts'), Path.home() / 'AppData/Local/Microsoft/Windows/Fonts', Path('/usr/share/fonts')):
        if (d / 'SamsungOneKorean-700.ttf').exists():
            L.FONTDIR = d
            L.SAMSUNG_FILE = {w: 'SamsungOneKorean-%d.ttf' % w for w in (400, 500, 700)}
            break

sys.path.insert(0, str(HERE.parent / '11_중간본'))
import build as HW  # noqa: E402
import samsung_fonts  # noqa: E402
import foreign as F  # noqa: E402
import apx_new as A  # noqa: E402
from hiw_lib import q  # noqa: E402
from pptx import Presentation  # noqa: E402

# make.py는 불러오지 않는다(불러올 때 How it Works의 흐린 그림을 out/로 복사하는데, 그 그림이 없는 컴퓨터에서도 돌게). 쓰는 것만 같은 값으로.
BASE = HERE.parent / '02_PPT' / 'Weave_결선_중간본.pptx'
WORK = HERE / 'out' / 'ondevice'
DECK = HERE / 'Weave_통합본_Claude.pptx'
AFTER = 'Q. 외부 앱 데이터 접근 권한과 사생활 보호는 어떻게 하나요?'     # 이 제목의 장 바로 뒤에 끼움


def one():
    WORK.mkdir(parents=True, exist_ok=True)
    prs = Presentation(str(BASE))                      # make.build, fx_try.frame과 같은 틀
    old = list(prs.slides)
    layout = old[5].slide_layout
    note_tpl = [copy.deepcopy(sh._element) for sh in old[5].notes_slide.placeholders]
    HW.add_slides(prs, layout, note_tpl, [A.s_ondevice()])
    drop_slides(prs, old)
    samsung_fonts.remap_prs(prs)
    dst = WORK / 'one.pptx'
    prs.save(str(dst))
    print('→', dst.relative_to(HERE))
    return dst


def drop_slides(prs, slides):
    """make.drop_slides와 같음."""
    parts = {s.part for s in slides}
    lst = prs.slides._sldIdLst
    for e in list(lst):
        if prs.part.related_part(e.get(q('r:id'))) in parts:
            prs.part.drop_rel(e.get(q('r:id')))
            lst.remove(e)


def title_of(slide):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            return sh.text_frame.text.strip()
    return ''


def splice(src, dst):
    prs = Presentation(str(DECK))
    slides = list(prs.slides)
    at = [i for i, s in enumerate(slides) if title_of(s).replace('\n', '') == AFTER]
    assert len(at) >= 1, '끼울 자리(%s)를 찾지 못함' % AFTER
    at = at[0]                                         # 이전 판 묶음에도 같은 제목이 있을 수 있으니 처음 것(Appendix)
    if any(title_of(s).startswith('Q. 왜 Weave는 온디바이스') for s in slides[:at + 3]):
        print('   이미 들어 있음 → 그대로 둠')
        return None
    note_tpl = [copy.deepcopy(sh._element) for sh in slides[at].notes_slide.placeholders]
    new = F.copy_slide(prs, list(Presentation(str(src)).slides)[0], note_tpl)
    lst = prs.slides._sldIdLst
    ids = list(lst)
    moved = ids[-1]
    lst.remove(moved)
    lst.insert(at + 1, moved)
    prs.save(str(dst))
    print('→ %s  (%d쪽으로 끼움, 모두 %d장)' % (dst.name, at + 2, len(ids)))
    return at + 2


def png(deck, page, out):
    """PowerPoint로 그 장 하나를 찍는다(창 없이 열고 닫음. 이미 열린 다른 덱은 건드리지 않음)."""
    ps = ("$pp = New-Object -ComObject PowerPoint.Application; "
          "$p = $pp.Presentations.Open('%s', $true, $false, $false); "
          "$p.Slides.Item(%d).Export('%s', 'PNG', 1920, 1080); $p.Close()" % (deck, page, out))
    subprocess.run(['powershell', '-NoProfile', '-Command', ps], check=True)
    print('→', out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--inplace', action='store_true')
    ap.add_argument('--png', action='store_true')
    a = ap.parse_args()
    src = one()
    dst = DECK if a.inplace else HERE / 'Weave_통합본_Claude_온디바이스.pptx'
    page = splice(src, dst)
    if a.png and page:
        png(str(dst), page, str(WORK / 'slide.png'))


if __name__ == '__main__':
    sys.exit(main())
