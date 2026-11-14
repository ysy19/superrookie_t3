#!/usr/bin/env python3
"""2쪽("잠깐만요, 찾아보고 다시 알려 드릴게요.")의 통화 소리 → assets/intro/call_intro.wav, 덱 2쪽에 넣기

  python call_sound.py            소리를 다시 만들고 Weave_통합본_Claude.pptx 2쪽에 넣는다(이미 들어 있으면 소리와 때만 바꿈)
  python call_sound.py --wav      소리만 다시 만든다

장이 열리면 저절로 나오는 차례(초, 소리가 시작한 때부터)
  0.10 ~ 1.10   통화 연결음 한 번 (440 + 480 Hz)
  1.90 ~ 2.45   둘째 연결음, 상대가 받으면서 끊김
  2.45          받는 소리. 이때 '통화 중' 표시가 뜬다 (PILL_MS)
  2.80 ~ 11.37  상대가 하는 말: "안녕하세요, 위브 프로님~!"(또렷하게) → 무슨 말인지 모르게 뭉갠 말 → "알려주실 수 있나요~?"(또렷하게)
  11.70         "잠깐만요, …" 문장이 한 글자씩 뜨기 시작 (QUOTE_MS)

연결음과 받는 소리는 여기서 합성하고, 상대의 말은 Supertonic 3(기기 안에서 도는 신경망 음성, 목소리 SAY_VOICE)으로 읽혀 만든다(받아 온 음원 없음).
가운데 토막은 뜻이 없는 말이다. Supertonic 환경(~/.cache/weave-tools)이 없으면 저장소에 있는 wav를 그대로 쓴다. make.py는 INTRO_PATCH['in02']에서 patch()를 부른다(wav는 저장소에 있는 것을 그대로 씀).
"""
import argparse
import copy
import os
import wave
from pathlib import Path

from lxml import etree

HERE = Path(__file__).resolve().parent
WAV = HERE / 'assets' / 'intro' / 'call_intro.wav'
DECK = HERE / 'Weave_통합본_Claude.pptx'

SR = 44100
RINGS = ((0.10, 1.10), (1.90, 2.45))   # 연결음이 울리는 구간(초). 둘째는 받는 순간에 끊김
PICK = 2.45                            # 상대가 받는 때
TALK = 2.80                            # 상대가 말을 시작하는 때
LEN = 11.60                             # 소리 길이
PILL_MS = round(PICK * 1000)           # '통화 중' 표시가 뜨는 때(ms, 장이 열린 뒤)
QUOTE_GAP = 100                        # 소리가 끝난 뒤 문장이 뜨기까지(ms)
QUOTE_MS = round(LEN * 1000) + QUOTE_GAP
QUOTE_WAS = 1000                       # 소리를 넣기 전에 문장이 뜨기 시작하던 때(ms). talk.py가 말을 떼는 때를 그만큼 늦춤
VOL = 0.8                              # 덱에서의 소리 크기(표지, 모션그래픽 영상과 같음)
NAME = 'Call sound'                    # 소리 도형의 이름. 화면 왼쪽 밖에 둔다(편집 화면에서만 보임)

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
RT_AUDIO = R_NS + '/audio'
RT_MEDIA = 'http://schemas.microsoft.com/office/2007/relationships/media'


def p(tag):
    return '{%s}%s' % (P_NS, tag)


# ------------------------------------------------------------------ 소리 합성
def _band(x, lo, hi, n=2):
    """lo ~ hi Hz만 남긴다(버터워스 n차의 크기 응답을 주파수 쪽에서 곱함: 위상은 그대로)."""
    import numpy as np
    f = np.fft.rfftfreq(len(x), 1 / SR)
    f[0] = 1e-9
    g = 1 / np.sqrt(1 + (lo / f) ** (2 * n)) / np.sqrt(1 + (f / hi) ** (2 * n))
    return np.fft.irfft(np.fft.rfft(x) * g, len(x))


def _edges(x, a, r):
    """앞 a초, 뒤 r초를 부드럽게."""
    import numpy as np
    na, nr = int(a * SR), int(r * SR)
    x = x.copy()
    if na:
        x[:na] *= 0.5 - 0.5 * np.cos(np.pi * np.arange(na) / na)
    if nr:
        x[-nr:] *= 0.5 + 0.5 * np.cos(np.pi * np.arange(nr) / nr)
    return x


# 상대가 하는 말: Supertonic 3(기기 안에서 도는 신경망 음성. How it Works 낭독과 같은 도구)으로 세 토막을 한 문장으로 이어 한 번에 읽혀
# 억양과 숨이 실제 사람처럼 이어지게 하고, 토막 사이의 쉼을 경계로 나눠 가운데 토막만 뭉갠다.
# 11-14에 받은 검토: 배음을 더한 합성음은 '소음 같다' → Windows 음성(Heami). '글자가 잘 들린다' → 지연 복사와 높은 소리 덜어 내기로 뭉갬.
# '인사와 끝의 부탁은 직장인 목소리로 정확하게' → 토막으로 나눔. '실제 사람처럼 자연스럽게, 뭉갠 곳이 울리지 않게'
# → 신경망 음성으로 바꾸고, 울림을 만들던 지연 복사(예전 SAY_SMEAR)를 빼고 높은 소리만 더 깊게 덜어 냄.
# 'M4로, 님~!과 있나요~?로, 뭉갠 말은 2초쯤 더 길게, 무슨 말인지 아예 모르게' → 물결표, 뜻 없는 말을 늘림, 조각 뒤집기(SAY_SCRAMBLE).
# 뭉갠 정도는 받아쓰기(Whisper large-v3-turbo, 기기 안)로 쟀다: 높은 소리만 덜어 내면(300~650 Hz) 음절 차례를 거의 맞히고,
# 조각 뒤집기를 더하면 말소리를 찾지 못한다('시청해주셔서 감사합니다' 같은 헛소리만 냄). 또렷한 두 마디는 글대로 받아씀.
# (읽을 글, 또렷하게 들리게 두는가). 가운데 토막은 아무 뜻이 없는 말
# '님~!'과 '있나요~?'의 물결표: Supertonic이 말끝을 높게 끌거나(님~!) 더 올려 끈다(있나요~?). 물결표 없이 읽힌 것보다 말끝이 0.05~0.1초 길고 70~80 Hz 높음
SAY = (('안녕하세요, 위브 프로님~!', True),
       ('소비카 누렌 다미 요르나 호르케 나불레, 미루 지나서 테사 바시케를,', False),
       ('알려주실 수 있나요~?', True))
SAY_VOICE = 'M4'                       # Supertonic 3 목소리(F1~F5, M1~M5). 11-14에 F5, F1, F3, M1, M4를 견줘 듣고 고름
SAY_SPEED = 1.10                       # 읽는 빠르기(1.05가 모델 기본). 통화하듯 조금 빠르게
SAY_STEPS = 12                         # 합성 단계 수(많을수록 고름)
SAY_CLEAR = (180, 3800)                # 또렷한 토막: 전화 너머 목소리의 대역(Hz). 완만하게 거름
SAY_SCRAMBLE = (0.09, 0.008)           # 뭉개는 토막: 이 길이(초)의 조각마다 앞뒤를 뒤집고 조각 사이만 짧게 겹쳐 이음(울림 없이). 음절의 차례가 흐트러져 낱말이 들리지 않음
SAY_BLUR = (80, 500, 6)                # 뭉개는 토막: 그다음 이 대역(Hz)만 남김(차수 6). 벽 너머처럼 높낮이와 박자만 들림
SAY_LEVEL = 0.8                        # 뭉개는 토막의 크기(또렷한 토막에 견줘)
SAY_PAUSE = 0.40                       # 토막 사이의 쉼을 이만큼까지만(초)
TOOLS = Path.home() / '.cache' / 'weave-tools'
TTS_PY = TOOLS / 'tts-venv' / ('Scripts/python.exe' if os.name == 'nt' else 'bin/python')
TTS_MODEL = TOOLS / 'supertonic-3'
TTS_WORKER = HERE.parent / '09_HowItWorks' / '장표5_claude' / 'tts_worker.py'


def _tts(texts):
    """Supertonic 3으로 읽은 소리들(float, 44.1 kHz). 같은 글은 늘 같은 소리. 환경이 없으면 ImportError(저장소의 wav를 그대로 씀)."""
    import subprocess
    import tempfile
    import numpy as np
    import hashlib
    cache = HERE / 'out' / 'call_tts'                           # 같은 글·목소리·빠르기는 한 번만 읽힘(저장소에는 올리지 않음)
    keys = [cache / ('%s.npy' % hashlib.sha1(('%s|%s|%s|%s' % (SAY_VOICE, SAY_SPEED, SAY_STEPS, s)).encode('utf-8')).hexdigest()[:16]) for s in texts]
    if all(k.exists() for k in keys):
        return [np.load(k) for k in keys]
    if not (TTS_PY.exists() and (TTS_MODEL / 'onnx' / 'vocoder.onnx').exists()):
        raise ImportError('Supertonic 3 환경이 없음(../09_HowItWorks/장표5_claude/README.md의 낭독 준비)')
    with tempfile.TemporaryDirectory() as d:
        lst = Path(d) / 'list.txt'
        lst.write_text(''.join('s%d\t%s\n' % (i, s) for i, s in enumerate(texts)), encoding='utf-8')
        env = dict(os.environ, HF_HUB_OFFLINE='1', HF_HUB_DISABLE_TELEMETRY='1')
        r = subprocess.run([str(TTS_PY), '-I', str(TTS_WORKER), 'synth', str(TTS_MODEL), str(lst), d, SAY_VOICE, str(SAY_SPEED), str(SAY_STEPS)],
                           capture_output=True, text=True, env=env)
        assert r.returncode == 0 and 'done %d' % len(texts) in r.stdout, r.stderr[-1500:]
        out = []
        for i in range(len(texts)):
            with wave.open(str(Path(d) / ('s%d.wav' % i))) as w:
                assert (w.getframerate(), w.getnchannels(), w.getsampwidth()) == (SR, 1, 2)
                out.append(np.frombuffer(w.readframes(w.getnframes()), '<i2') / 32768)
    cache.mkdir(parents=True, exist_ok=True)
    for k, x in zip(keys, out):
        np.save(k, x)
    return out


def _voiced(x):
    """10ms마다 소리가 나는가(가장 큰 소리의 2%를 넘는가)."""
    import numpy as np
    n = int(0.01 * SR)
    return np.sqrt(np.convolve(x ** 2, np.ones(n) / n, 'same')) > 0.02 * np.abs(x).max()


def _trim(x):
    import numpy as np
    on = np.flatnonzero(_voiced(x))
    return x[max(on[0] - int(0.01 * SR), 0):on[-1] + int(0.04 * SR)]


def _reverse(x, win, fade):
    """짧은 조각(win초)마다 앞뒤를 뒤집고, 조각 사이만 fade초 겹쳐 잇는다. 반씩 겹쳐 이으면 두 소리가 겹쳐 울려서 겹침을 짧게 둠."""
    import numpy as np
    n, f = int(win * SR), int(fade * SR)
    ramp = np.linspace(0, 1, f)
    out = np.zeros(len(x) + n)
    for i in range(0, len(x), n - f):
        c = x[i:i + n][::-1].copy()
        if len(c) > 2 * f:
            c[:f] *= ramp
            c[-f:] *= ramp[::-1]
        out[i:i + len(c)] += c
    return out[:len(x)]


def _speech(pieces=None):
    """상대가 하는 말: SAY를 한 문장으로 읽힌 뒤, 토막 사이의 쉼에서 나눠 뭉갤 토막만 뭉갠다.
    나눌 자리는 토막을 따로 읽힌 길이로 어림한다. 앞쪽 경계는 앞에서부터(첫 토막의 길이), 뒤쪽 경계는 뒤에서부터(끝 토막의 길이) 재어
    그 둘레 0.7초 안에서 가장 긴 쉼(가운데 토막 안의 쉼에 끌려가지 않게). pieces에 목록을 주면 손본 토막을 차례로 넣어 준다(점검용)."""
    import numpy as np
    whole, *parts = _tts([' '.join(s for s, _ in SAY)] + [s for s, _ in SAY])
    x = _trim(whole)
    lens = [len(_trim(p)) for p in parts]
    on = _voiced(x)
    d = np.diff(on.astype(int))
    falls, rises = np.flatnonzero(d < 0) + 1, np.flatnonzero(d > 0) + 1
    gaps = [(s, rises[rises > s][0]) for s in falls if (rises > s).any()]                # 소리가 꺼졌다 다시 켜지는 자리(쉼)
    gaps = [(s, e) for s, e in gaps if e - s > 0.06 * SR]
    cuts = []
    for k in range(1, len(SAY)):
        front = k < len(SAY) / 2
        want = sum(lens[:k]) if front else len(x) - sum(lens[k:])
        near = [g for g in gaps if abs((g[0] if front else g[1]) - want) < 0.7 * SR]      # 문장 속에서는 말끝을 더 끌기도 해서 넉넉히
        assert near, '토막 사이의 쉼을 찾지 못함: %d번째' % k
        cuts.append(max(near, key=lambda g: g[1] - g[0]))                                  # 그 둘레에서 가장 긴 쉼
    segs, at = [], 0
    for s, e in cuts:                                           # 쉼은 SAY_PAUSE까지만 남기고 그 가운데에서 나눔
        keep = min(e - s, int(SAY_PAUSE * SR)) // 2
        segs.append(x[at:s + keep])
        at = e - keep
    segs.append(x[at:])
    ref = None
    out = []
    for seg, (_, clear) in zip(segs, SAY):
        seg = _band(seg, *SAY_CLEAR) if clear else _band(_reverse(seg, *SAY_SCRAMBLE), *SAY_BLUR)
        rms = np.sqrt(np.mean(seg[_voiced(seg)] ** 2))
        ref = ref or (rms if clear else None)
        out.append(_edges(seg, 0.01, 0.01) * (1 if clear else SAY_LEVEL * (ref or rms) / rms))
    if pieces is not None:
        pieces.extend(out)
    y = np.concatenate(out)
    return y / np.abs(y).max()


def synth(seed=7):
    """통화 소리 전체(float, -1 ~ 1)."""
    import numpy as np
    rng = np.random.default_rng(seed)
    n = int(LEN * SR)
    x = np.zeros(n)
    for a, b in RINGS:                                          # 통화 연결음
        t = np.arange(int((b - a) * SR)) / SR
        ring = 0.15 * (np.sin(2 * np.pi * 440 * t) + np.sin(2 * np.pi * 480 * t))
        x[int(a * SR):int(a * SR) + len(t)] += _edges(ring, 0.015, 0.004 if b == PICK else 0.015)
    i = int(PICK * SR)                                          # 받는 소리: 짧은 딸깍 둘(선로 잡음은 깔지 않음: 말소리가 깨끗하게 들리게)
    click = np.zeros(int(0.1 * SR))
    for at, amp, d in ((0.0, 0.5, 0.012), (0.055, 0.22, 0.008)):
        m = int(d * SR)
        j = int(at * SR)
        click[j:j + m] += amp * rng.standard_normal(m) * np.exp(-np.arange(m) / (m / 4))
    x[i:i + len(click)] += _band(click, 350, 3000)
    sp = _speech()                                              # 상대의 말
    j = int(TALK * SR)
    assert j + len(sp) <= n - int(0.1 * SR), '말이 소리 길이를 넘음: %.2f초에 끝남' % ((j + len(sp)) / SR)
    x[j:j + len(sp)] += 0.7 * sp
    return _edges(0.8 * x / np.abs(x).max(), 0.0, 0.05)


def write_wav(path=WAV):
    import numpy as np
    x = synth()
    path.parent.mkdir(parents=True, exist_ok=True)
    with wave.open(str(path), 'wb') as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(SR)
        w.writeframes((x * 32767).astype('<i2').tobytes())
    return len(x) / SR


# ------------------------------------------------------------------ 덱에 넣기
PIC_XML = ('<p:pic xmlns:p="%s" xmlns:a="%s" xmlns:r="%s"><p:nvPicPr><p:cNvPr id="%%d" name="%s"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>'
           '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:audioFile r:link="%%s"/><p:extLst>'
           '<p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="%s" r:embed="%%s"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>'
           '<p:blipFill><a:blip r:embed="%%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
           '<p:spPr><a:xfrm><a:off x="-533400" y="152400"/><a:ext cx="304800" cy="304800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
           '</p:pic>') % (P_NS, A_NS, R_NS, NAME, P14)
PLAY_XML = ('<p:par xmlns:p="%s"><p:cTn id="%%d" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="withEffect">'
            '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr>'
            '<p:cTn id="%%d" dur="%%d" fill="hold"/><p:tgtEl><p:spTgt spid="%%d"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>') % P_NS
NODE_XML = ('<p:audio xmlns:p="%s"><p:cMediaNode vol="%%d" showWhenStopped="0"><p:cTn id="%%d" fill="hold" display="0"><p:stCondLst>'
            '<p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="%%d"/></p:tgtEl></p:cMediaNode></p:audio>') % P_NS


def _shape(tree, name):
    hit = [e.getparent().getparent() for e in tree.iter(p('cNvPr')) if e.get('name') == name]
    assert len(hit) <= 1, (name, len(hit))
    return hit[0] if hit else None


def _effect(sl, spid):
    """spid 도형에 걸린 효과의 cTn(presetID가 있는 것)."""
    hit = [c for c in sl._element.iter(p('cTn')) if c.get('presetID') and {t.get('spid') for t in c.iter(p('spTgt'))} == {spid}]
    assert len(hit) == 1, (spid, len(hit))
    return hit[0]


def _delay(ctn, ms):
    ctn.find(p('stCondLst')).find(p('cond')).set('delay', str(ms))


def patch(sl):
    """2쪽: 통화 소리를 넣고 장이 열리면 저절로 나오게 한다. '통화 중' 표시는 상대가 받는 때에, 문장은 상대의 말이 끝난 뒤에 뜬다.
    이미 넣은 장이면 소리와 때만 지금 값으로 바꾼다."""
    from pptx.media import Video
    tree = sl._element.find(p('cSld')).find(p('spTree'))
    pill, quote = _shape(tree, '!!call'), _shape(tree, '!!quote')
    assert pill is not None and quote is not None, '2쪽이 아님'
    pid = pill.find(p('nvGrpSpPr')).find(p('cNvPr')).get('id')
    qid = quote.find(p('nvSpPr')).find(p('cNvPr')).get('id')
    r = lambda k: '{%s}%s' % (R_NS, k)
    pic = _shape(tree, NAME)
    if pic is None:
        part = sl.part._package.get_or_add_media_part(Video.from_path_or_file_like(str(WAV), 'audio/x-wav'))
        rid_m, rid_a = sl.part.relate_to(part, RT_MEDIA), sl.part.relate_to(part, RT_AUDIO)
        icon = _shape(tree, 'Call icon').find('.//{%s}blip' % A_NS).get(r('embed'))        # 소리 도형의 그림은 '통화 중' 표시의 전화 아이콘을 같이 씀
        sid = max(int(e.get('id')) for e in tree.iter(p('cNvPr')) if (e.get('id') or '').isdigit()) + 1
        tree.append(etree.fromstring(PIC_XML % (sid, rid_a, rid_m, icon)))
        ids = [int(c.get('id')) for c in sl._element.iter(p('cTn')) if (c.get('id') or '').isdigit()]
        host = _effect(sl, pid).getparent().getparent()                                     # '통화 중' 표시의 효과가 든 칸(장이 열리면 바로 시작)
        host.insert(0, etree.fromstring(PLAY_XML % (max(ids) + 1, max(ids) + 2, round(LEN * 1000), sid)))
        root = sl._element.find(p('timing')).find(p('tnLst')).find(p('par')).find(p('cTn')).find(p('childTnLst'))
        root.append(etree.fromstring(NODE_XML % (round(VOL * 100000), max(ids) + 3, sid)))
    else:
        sid = int(pic.find(p('nvPicPr')).find(p('cNvPr')).get('id'))
        ext = pic.find(p('nvPicPr')).find(p('nvPr')).find(p('extLst'))
        sl.part.related_part(ext.find('.//{%s}media' % P14).get(r('embed')))._blob = WAV.read_bytes()
        for c in _effect(sl, str(sid)).iter(p('cTn')):
            if (c.get('dur') or '').isdigit():
                c.set('dur', str(round(LEN * 1000)))
    _delay(_effect(sl, pid), PILL_MS)
    q_eff = _effect(sl, qid)
    _delay(q_eff, QUOTE_GAP)
    _delay(q_eff.getparent().getparent().getparent(), QUOTE_MS - QUOTE_GAP)                 # 문장이 든 칸은 앞 칸(소리)이 끝나는 때에 시작
    return sid


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--wav', action='store_true', help='소리만 다시 만든다')
    a = ap.parse_args()
    try:
        print('→', WAV.name, '%.2f초' % write_wav())
    except ImportError as e:                                    # Supertonic 환경이 없는 곳: 저장소에 있는 소리를 그대로 씀
        print('   ! %s → %s는 그대로 둠' % (e, WAV.name))
    if a.wav:
        return
    from pptx import Presentation
    prs = Presentation(str(DECK))
    patch(prs.slides[1])
    tmp = DECK.with_suffix('.tmp.pptx')
    prs.save(str(tmp))
    tmp.replace(DECK)
    print('→', DECK.name, '2쪽: 표시 %.2f초, 문장 %.2f초' % (PILL_MS / 1000, QUOTE_MS / 1000))


if __name__ == '__main__':
    main()
