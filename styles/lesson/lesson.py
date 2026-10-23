"""Видео-урок: горизонтальное видео 1920×1080 (YouTube, курс, интенсив).

Сцена – карточка во весь кадр: метка, заголовок капсом с акцентным словом, ниже список / сравнение / цифра / цепочка /
скриншот. Лицо всегда в кадре: окно 420×420 справа внизу или на весь кадр в сценах face. Правила – STYLE.md.

    python3 lesson.py plan.json                 – рендер
    python3 lesson.py plan.json --frames 1,40,90 – контрольные кадры в check/ (с лицом, как в ролике)
    python3 lesson.py plan.json --sheet          – лист всех карточек check/lesson_sheet.jpg

plan.json: source, transcript, keep или keep_file, face_crop [x, y, сторона] – квадрат лица для окна,
full_crop [x, y, w, h] – кадр 16:9 для сцен face, fix, out, scenes.
Сцена: from (секунды ролика или слово, "слово#2" – второе вхождение), kind, whoosh (true – звук на смене блока).
kind: face | img | imgs | list | compare | big | chain | word. Поля – STYLE.md, раздел 4.
Картинки (src, srcs, side) – пути от папки плана."""
import json
import os
import re
import subprocess
import sys
import urllib.parse
from pathlib import Path

# запуск из любого python: переключиться на окружение .venv папки montage (его создаёт install.py)
_VENV = Path(__file__).resolve().parents[2] / ".venv"
_VPY = _VENV / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
if _VPY.exists() and Path(sys.prefix).resolve() != _VENV.resolve():
    sys.exit(subprocess.run([str(_VPY), *sys.argv]).returncode)

from PIL import Image, ImageDraw  # noqa: E402

KIT = Path(__file__).resolve().parent
W, H, FPS = 1920, 1080, 30
PIP = (W - 56 - 420, H - 56 - 420, 420)   # окно лица: x, y, сторона
RADIUS = 36
ZOOM = 0.06                               # наезд в сценах face: 1.0 → 1.06 за сцену
WHOOSH_GAIN = 0.315


def run(*a):
    subprocess.run(["ffmpeg", "-y", "-v", "error", *map(str, a)], check=True)


def load_words(plan, root):
    d = json.load(open(root / plan["transcript"], encoding="utf-8"))
    if "keep" not in plan:
        plan["keep"] = json.load(open(root / plan["keep_file"]))
    fix = plan.get("fix", {})
    words, off = [], 0.0
    for a, b in plan["keep"]:
        for w in d["words"]:
            if a <= w["start"] < b:
                t = w.get("text", w.get("word", "")).strip()
                words.append({"t": fix.get(t, t), "o": t, "s": round(w["start"] - a + off, 3)})
        off += b - a
    return words, off


def norm(s):
    return re.sub(r"[^\wё-]", "", s.lower())


def resolve(at, words, t0):
    if isinstance(at, (int, float)):
        return float(at)
    key, n = (at.split("#") + ["1"])[:2]
    key, n = norm(key), int(n)
    for w in words:
        if w["s"] >= t0 - 0.05 and (norm(w["t"]).startswith(key) or norm(w["o"]).startswith(key)):
            n -= 1
            if not n:
                return w["s"]
    raise ValueError(f"слово «{at}» не найдено после {t0:.2f}")


def prepare(plan_path):
    root = plan_path.parent
    plan = json.load(open(plan_path, encoding="utf-8"))
    words, total = load_words(plan, root)
    sc = plan["scenes"]
    t0 = 0.0
    for s in sc:   # якорь ищется после начала предыдущей сцены: «слово» не схватит то же слово раньше
        s["from"] = t0 = resolve(s["from"], words, t0)
    for i, s in enumerate(sc):
        s["to"] = sc[i + 1]["from"] if i + 1 < len(sc) else total
        if s["to"] <= s["from"]:
            raise ValueError(f"сцена {i} ({s.get('title', s['kind'])}) начинается не позже следующей: {s['from']:.2f}")
    if sc and sc[0]["from"] > 0.05:
        raise ValueError("первая сцена должна начинаться с 0")
    return root, plan, total


def uri(root, p):
    return (root / p).resolve().as_uri()


def frames(root, plan, only=None):
    """Карточки сцен → root/lesson_frames/NN.png (кроме face)."""
    from patchright.sync_api import sync_playwright
    out = root / "lesson_frames"
    out.mkdir(exist_ok=True)
    with sync_playwright() as pw:
        try:
            b = pw.chromium.launch(channel="chrome", headless=True, args=["--allow-file-access-from-files"])
        except Exception:
            b = pw.chromium.launch(headless=True, args=["--allow-file-access-from-files"])
        pg = b.new_page(viewport={"width": W, "height": H})
        for i, s in enumerate(plan["scenes"]):
            if s["kind"] == "face" or (only is not None and i not in only):
                continue
            spec = {k: v for k, v in s.items() if k not in ("from", "to", "whoosh")}
            for k in ("src", "side"):
                if k in spec:
                    spec[k] = uri(root, spec[k])
            if "srcs" in spec:
                spec["srcs"] = [uri(root, x) for x in spec["srcs"]]
            pg.goto(KIT.joinpath("frame.html").as_uri() + "#" + urllib.parse.quote(json.dumps(spec, ensure_ascii=False)))
            pg.reload(wait_until="networkidle")
            pg.evaluate("document.fonts.ready")
            pg.wait_for_timeout(300)
            pg.screenshot(path=str(out / f"{i:02d}.png"))
        b.close()
    return out


def mask(path):
    m = Image.new("L", (PIP[2], PIP[2]), 0)
    ImageDraw.Draw(m).rounded_rectangle((0, 0, PIP[2] - 1, PIP[2] - 1), RADIUS, fill=255)
    m.save(path)


def to_src(plan, t):
    c = 0.0
    for a, b in plan["keep"]:
        if t <= c + b - a:
            return a + t - c
        c += b - a
    return plan["keep"][-1][1]


def check(root, plan, ts):
    """Контрольные кадры: карточка + лицо из исходника, как в готовом ролике."""
    sc = plan["scenes"]
    idx = [next(i for i, s in enumerate(sc) if s["from"] <= t < s["to"] or i == len(sc) - 1) for t in ts]
    fr = frames(root, plan, set(idx))
    (root / "check").mkdir(exist_ok=True)
    for t, i in zip(ts, idx):
        tmp = root / "check" / "_src.png"
        run("-ss", f"{to_src(plan, t):.3f}", "-i", root / plan["source"], "-frames:v", "1", tmp)
        src = Image.open(tmp).convert("RGB")
        if sc[i]["kind"] == "face":
            x, y, w, h = plan["full_crop"]
            im = src.crop((x, y, x + w, y + h)).resize((W, H))
        else:
            x, y, d = plan["face_crop"]
            im = Image.open(fr / f"{i:02d}.png").convert("RGB")
            m = Image.new("L", (PIP[2], PIP[2]), 0)
            ImageDraw.Draw(m).rounded_rectangle((0, 0, PIP[2] - 1, PIP[2] - 1), RADIUS, fill=255)
            im.paste(src.crop((x, y, x + d, y + d)).resize((PIP[2], PIP[2])), PIP[:2], m)
        out = root / "check" / f"lesson_{t:07.2f}.jpg"
        im.save(out, quality=85)
        tmp.unlink()
        print(out)


def sheet(root, plan):
    fr = frames(root, plan)
    ims = [(i, Image.open(fr / f"{i:02d}.png").resize((480, 270))) for i, s in enumerate(plan["scenes"]) if s["kind"] != "face"]
    cols = 5
    sh = Image.new("RGB", (cols * 490, max(1, -(-len(ims) // cols)) * 300), "#444")
    d = ImageDraw.Draw(sh)
    for k, (i, im) in enumerate(ims):
        x, y = (k % cols) * 490, (k // cols) * 300
        sh.paste(im, (x, y))
        d.text((x + 4, y + 274), f"{i:02d}  {plan['scenes'][i]['from']:.1f} с", fill="white")
    (root / "check").mkdir(exist_ok=True)
    sh.save(root / "check" / "lesson_sheet.jpg", quality=85)
    print(root / "check" / "lesson_sheet.jpg")


def whoosh_track(path, events, total):
    """Одиночный «вжух» на смене крупного блока: шум 900±700 Гц, огибающая sin^1.5, 0,38 с, старт за 0,1 с до смены."""
    import wave
    import numpy as np
    sr = 48000
    out = np.zeros(int(total * sr) + 1, np.float32)
    L = int(0.38 * sr)
    rng = np.random.default_rng(7)
    for ev in events:
        sp = np.fft.rfft(rng.standard_normal(L))
        fq = np.fft.rfftfreq(L, 1 / sr)
        sp *= np.exp(-0.5 * ((fq - 900) / 700) ** 2)
        x = np.fft.irfft(sp, L)
        x /= np.abs(x).max() + 1e-9
        env = np.sin(np.linspace(0, np.pi, L)) ** 1.5
        s0 = max(0, int((ev - 0.10) * sr))
        e0 = min(len(out), s0 + L)
        out[s0:e0] += (x * env * WHOOSH_GAIN)[:e0 - s0]
    with wave.open(str(path), "wb") as w:
        w.setnchannels(1); w.setsampwidth(2); w.setframerate(sr)
        w.writeframes((np.clip(out, -1, 1) * 32767).astype(np.int16).tobytes())


def render(root, plan, total):
    parts = root / "lesson_parts"
    parts.mkdir(exist_ok=True)
    fr = frames(root, plan)
    # 1. склейка без дублей
    K = plan["keep"]
    f = ";".join(f"[0:v]trim={a}:{b},setpts=PTS-STARTPTS,fps={FPS}[v{i}];[0:a]atrim={a}:{b},asetpts=PTS-STARTPTS,"
                 f"afade=t=in:d=0.02,afade=t=out:st={b - a - 0.02}:d=0.02[a{i}]" for i, (a, b) in enumerate(K))
    f += ";" + "".join(f"[v{i}][a{i}]" for i in range(len(K))) + f"concat=n={len(K)}:v=1:a=1[v][a]"
    base = parts / "base.mov"
    run("-i", root / plan["source"], "-filter_complex", f, "-map", "[v]", "-map", "[a]",
        "-c:v", "libx264", "-crf", "16", "-preset", "fast", "-c:a", "pcm_s16le", base)
    mask(parts / "mask.png")
    # 2. сцены: границы по сетке кадров, длина куска ровно в кадрах – иначе склейка уезжает от звука
    sc = plan["scenes"]
    edges = [round(s["from"] * FPS) for s in sc] + [round(total * FPS)]
    lst = []
    for i, s in enumerate(sc):
        n = edges[i + 1] - edges[i]
        if n < 1:
            continue
        out = parts / f"{i:02d}.mp4"
        common = ["-ss", f"{edges[i] / FPS:.4f}", "-i", base]
        enc = ["-frames:v", n, "-an", "-c:v", "libx264", "-crf", "17", "-preset", "fast", "-pix_fmt", "yuv420p", "-r", FPS, out]
        if s["kind"] == "face":
            x, y, w, h = plan["full_crop"]
            run(*common, "-filter_complex",
                f"[0:v]crop={w}:{h}:{x}:{y},scale={W * 2}:{H * 2},"
                f"zoompan=z='1+{ZOOM}*on/{n}':d=1:s={W}x{H}:fps={FPS}:x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)'[v]",
                "-map", "[v]", *enc)
        else:
            x, y, d = plan["face_crop"]
            run(*common, "-loop", "1", "-i", fr / f"{i:02d}.png", "-i", parts / "mask.png", "-filter_complex",
                f"[0:v]crop={d}:{d}:{x}:{y},scale={PIP[2]}:{PIP[2]},setsar=1,format=rgba[f];[2:v]format=gray[m];[f][m]alphamerge[fm];"
                f"[1:v]fps={FPS},format=yuv420p[bg];[bg][fm]overlay={PIP[0]}:{PIP[1]}[v]",
                "-map", "[v]", *enc)
        lst.append(out)
    (parts / "list.txt").write_text("".join(f"file '{p.as_posix()}'\n" for p in lst))
    run("-f", "concat", "-safe", "0", "-i", parts / "list.txt", "-c", "copy", parts / "video.mp4")
    # 3. звук: голос −14 LUFS + «вжух» только на сценах с whoosh: true
    whoosh_track(parts / "whoosh.wav", [edges[i] / FPS for i, s in enumerate(sc) if s.get("whoosh")], total)
    run("-i", base, "-i", parts / "whoosh.wav", "-filter_complex",
        "[0:a]loudnorm=I=-14:TP=-1.5:LRA=11[v];[1:a]aresample=48000[w];[v][w]amix=inputs=2:normalize=0:duration=first[a]",
        "-map", "[a]", "-c:a", "pcm_s16le", parts / "audio.wav")
    out = root / plan.get("out", "final.mp4")
    run("-i", parts / "video.mp4", "-i", parts / "audio.wav", "-map", "0:v", "-map", "1:a", "-c:v", "copy",
        "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", "-shortest", out)
    print(out, f"{total:.1f} с")


def main():
    root, plan, total = prepare(Path(sys.argv[1]).resolve())
    for k in ("face_crop", "full_crop"):
        if k not in plan:
            sys.exit(f"В plan.json нет {k}: подобрать по tools/frames.py --grid (STYLE.md, раздел 3)")
    if "--frames" in sys.argv:
        check(root, plan, [float(x) for x in sys.argv[sys.argv.index("--frames") + 1].split(",")])
    elif "--sheet" in sys.argv:
        sheet(root, plan)
    else:
        render(root, plan, total)


if __name__ == "__main__":
    main()
