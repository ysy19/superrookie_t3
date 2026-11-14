#!/usr/bin/env python3
"""Appendix '왜 온디바이스인가' 장(apx_new.s_ondevice)의 선 그림 아이콘을 만든다.

  python tools/ondevice_assets.py   → assets/apx/ic-phone.png, ic-cloud.png, ic-lock.png, ic-shield.png (256×256, 투명 바탕)

아웃트로 덱의 다른 아이콘(assets/outro/ic-*.png)과 같은 규격: 24칸 격자에 굵기 2의 선 그림, 끝은 둥글게.
SVG를 PyMuPDF로 찍는다(Chromium 없이 Windows에서도 돌게).
"""
import sys
from pathlib import Path

import fitz  # PyMuPDF

HERE = Path(__file__).resolve().parent.parent
DST = HERE / 'assets' / 'apx'
GRAY, PURPLE, WHITE = '#8E88A0', '#6B4EFF', '#FFFFFF'

ICONS = {
    'ic-phone': (GRAY, '<rect x="6" y="2" width="12" height="20" rx="2.5"/><path d="M11 18h2"/>'),
    'ic-cloud': (GRAY, '<path d="M17.5 19H9a7 7 0 1 1 6.71-9h1.79a4.5 4.5 0 1 1 0 9Z"/>'),
    'ic-lock': (PURPLE, '<rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/><path d="M12 15v2"/>'),
    'ic-shield': (WHITE, '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/>'),
}


def svg(color, body, px=256):
    return ('<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 24 24" fill="none" stroke="%s" '
            'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">%s</svg>' % (px, px, color, body))


def main():
    DST.mkdir(parents=True, exist_ok=True)
    for name, (color, body) in ICONS.items():
        doc = fitz.open(stream=svg(color, body).encode(), filetype='svg')
        doc[0].get_pixmap(alpha=True).save(str(DST / (name + '.png')))
        print('→ assets/apx/%s.png' % name)


if __name__ == '__main__':
    sys.exit(main())
