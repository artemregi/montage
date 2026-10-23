"""Презентация для видео-урока: слайды картинками + черновик «какой слайд на какие слова».

    python3 tools/slides.py <проект> <презентация.pdf|.pptx|.key>   – слайды → slides/NN.png, текст → slides.json,
                                                                      лист check/slides_sheet.jpg
    python3 tools/slides.py <проект> --match [--start <слово|сек>]   – после plan.py: сопоставить слайды с речью →
                                                                      slides_match.md (таблица с флагами) и slides_draft.json

Сопоставление: каждая фраза ролика получает слайд. Слайды идут по порядку (назад не возвращаются, пропуск – штраф),
сходство – общие слова фразы и слайда (редкие слова весят больше, цифры – сильнее всего), плюс поправка на место
в ролике: при длине урока L и N слайдах слайд k ожидается около start + (k − 0,5)·(L − start)/N.
До --start (вступление без слайдов) – лицо. Черновик не финал: решение по каждому слайду – STYLE.md lesson, раздел 8."""
import argparse
import json
import math
import re
import shutil
import subprocess
import sys
from pathlib import Path

from common import fmt, load, project, save

STOP = set("""это этот эта эти того тому том тот такой такая такие так также тоже когда если чтобы потому поэтому
который которая которые которых какой какая какие какую можно нужно надо будет было были есть очень просто
своей свой свою свои твой твоя твои тебе тебя меня мной мене нами нему него неё ними всех всего всем всё всеми
сейчас теперь здесь тогда потом даже только ещё уже вообще между через после перед почему зачем сколько много
мало более менее самый самая самое сама сами себя себе один одна одно одни раз два три вот ведь например""".split())


def to_pdf(src, out_dir):
    """pptx / key / odp → pdf: LibreOffice, Keynote (Mac) или PowerPoint (Windows). Иначе – попросить PDF у автора."""
    pdf = out_dir / (src.stem + ".pdf")
    soffice = shutil.which("soffice") or shutil.which("libreoffice") or \
        next((p for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
                          "C:/Program Files/LibreOffice/program/soffice.exe") if Path(p).exists()), None)
    if src.suffix.lower() == ".key" or (sys.platform == "darwin" and not soffice and Path("/Applications/Keynote.app").exists()):
        subprocess.run(["osascript", "-e", f'tell application "Keynote"\nset d to open POSIX file "{src}"\n'
                        f'export d to POSIX file "{pdf}" as PDF\nclose d saving no\nend tell'], check=False)
    elif soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(src)], check=False)
    elif sys.platform == "win32":
        ps = (f"$p=New-Object -ComObject PowerPoint.Application;$d=$p.Presentations.Open('{src}',$true,$false,$false);"
              f"$d.SaveAs('{pdf}',32);$d.Close();$p.Quit()")
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=False)
    if not pdf.exists():
        sys.exit("Не получилось перевести презентацию в PDF. Попросить автора: «Файл → Экспорт → PDF» "
                 "(Google Slides: «Файл → Скачать → PDF») и дать путь к PDF.")
    return pdf


def extract(p, src):
    import pymupdf
    src = Path(src).expanduser().resolve()
    d = p / "slides"
    d.mkdir(exist_ok=True)
    for old in d.glob("*.png"):
        old.unlink()
    pdf = src if src.suffix.lower() == ".pdf" else to_pdf(src, d)
    doc = pymupdf.open(pdf)
    res = []
    for i, page in enumerate(doc, 1):
        zoom = 1920 / page.rect.width
        page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom)).save(d / f"{i:02d}.png")
        blocks = sorted((b for b in page.get_text("blocks") if b[4].strip()), key=lambda b: (round(b[1] / 20), b[0]))
        text = " ".join(" ".join(b[4].split()) for b in blocks)
        # заголовок – самый крупный текст слайда
        spans = [s for b in page.get_text("dict")["blocks"] for ln in b.get("lines", []) for s in ln["spans"] if s["text"].strip()]
        top = max((s["size"] for s in spans), default=0)
        title = " ".join(s["text"].strip() for s in spans if s["size"] >= top * 0.9)
        res.append({"n": i, "img": f"slides/{i:02d}.png", "title": title, "text": text})
    save(p, "slides.json", {"source": str(src), "slides": res})
    sheet(p, res)
    empty = [s["n"] for s in res if len(s["text"]) < 15]
    print(f"{len(res)} слайдов → {d}, текст → {p / 'slides.json'}")
    if empty:
        print(f"Почти без текста (картинка, схема): {', '.join(map(str, empty))} – их место определить только по смыслу, глазами")


def sheet(p, res):
    from PIL import Image, ImageDraw
    from frames import font
    cols, w = 4, 480
    ims = [Image.open(p / s["img"]).convert("RGB") for s in res]
    ims = [im.resize((w, round(im.height * w / im.width))) for im in ims]
    h = max(im.height for im in ims) + 36
    sh = Image.new("RGB", (cols * (w + 10), -(-len(ims) // cols) * h), "#333")
    dr = ImageDraw.Draw(sh)
    for k, im in enumerate(ims):
        x, y = (k % cols) * (w + 10), (k // cols) * h
        sh.paste(im, (x, y))
        dr.text((x + 6, y + im.height + 4), f"слайд {k + 1}", fill="white", font=font(24))
    (p / "check").mkdir(exist_ok=True)
    sh.save(p / "check" / "slides_sheet.jpg", quality=85)
    print(p / "check" / "slides_sheet.jpg")


def toks(t):
    out = []
    for w in re.findall(r"[\wё]+", t.lower().replace("ё", "е")):
        if w.isdigit():
            out.append("#" + w)
        elif len(w) >= 4 and w not in STOP:
            out.append(w[:5])
    return out


def phrases(p):
    """Фразы ролика: (начало, конец в ролике, текст) – как в phrases.md."""
    keep, words = load(p, "keep.json"), load(p, "transcript_src.json")["words"]
    res, off = [], 0.0
    for a, b in keep:
        res += [(w["start"] - a + off, min(w["end"], b) - a + off, w["text"]) for w in words if a <= w["start"] < b]
        off += b - a
    out, cur = [], []
    for w in res:
        cur.append(w)
        if w[2].rstrip().endswith((".", "!", "?", "…")) or len(cur) >= 12:
            out.append((cur[0][0], cur[-1][1], " ".join(x[2] for x in cur)))
            cur = []
    if cur:
        out.append((cur[0][0], cur[-1][1], " ".join(x[2] for x in cur)))
    return out, off


def match(p, start):
    sl = load(p, "slides.json")["slides"]
    ph, total = phrases(p)
    N, M = len(sl), len(ph)
    if isinstance(start, str):   # слово → время первой фразы, где оно есть
        key = start.lower()
        start = next((a for a, _, t in ph if key in t.lower()), None)
        if start is None:
            sys.exit(f"слово «{key}» не найдено в речи")
    s0 = next((j for j, (a, _, _) in enumerate(ph) if a >= (start or 0) - 0.05), 0)
    st = [set(toks(s["title"] + " " + s["text"])) for s in sl]
    df = {}
    for s in st:
        for t in s:
            df[t] = df.get(t, 0) + 1
    wt = lambda t: (2.0 if t.startswith("#") else 1.0) * math.log(1 + N / df[t])
    sim = [[0.0] * M for _ in range(N)]
    hits = [[set() for _ in range(M)] for _ in range(N)]
    for j, (_, _, t) in enumerate(ph):
        pt = set(toks(t))
        for i in range(N):
            common = pt & st[i]
            hits[i][j] = common
            sim[i][j] = sum(wt(x) for x in common) / math.sqrt(len(pt) + 1)
    top = max((v for r in sim for v in r), default=0) or 1.0
    span = max(1, M - s0)
    POS, STAY, NEXT, SKIP = 0.25, 0.0, 0.15, 0.6

    def score(i, j):   # i – номер слайда с 0; сходство 0..1 + место в ролике
        exp = (i + 0.5) / N
        return sim[i][j] / top + POS * (1 - min(1.0, abs((j - s0 + 0.5) / span - exp) * N / 2))

    # динамика: фразы до s0 – вступление (лицо), дальше слайды только вперёд
    NEG = -1e9
    best = [[NEG] * N for _ in range(M)]
    back = [[0] * N for _ in range(M)]
    for j in range(s0, M):
        for i in range(N):
            if j == s0:
                cand = [(-SKIP * i, -1)]
            else:
                cand = [(best[j - 1][k] - (STAY if k == i else NEXT + SKIP * (i - k - 1)), k) for k in range(i + 1)]
            v, k = max(cand)
            best[j][i], back[j][i] = v + score(i, j), k
    path = [None] * M
    if M > s0:
        i = max(range(N), key=lambda x: best[M - 1][x] - SKIP * (N - 1 - x))
        for j in range(M - 1, s0 - 1, -1):
            path[j] = i
            i = back[j][i]

    # отрезки слайдов
    segs = []
    for j in range(M):
        i = path[j]
        if segs and segs[-1]["i"] == i:
            segs[-1]["j"].append(j)
        else:
            segs.append({"i": i, "j": [j]})
    draft, rows, shown = [], [], set()
    for k, sg in enumerate(segs):
        a = 0.0 if k == 0 else ph[sg["j"][0]][0]
        b = ph[segs[k + 1]["j"][0]][0] if k + 1 < len(segs) else total
        i = sg["i"]
        if i is None:
            draft.append({"from": round(a, 2), "kind": "face", "_note": "вступление до первого слайда"})
            rows.append(f"| – | {fmt(a)} | {b - a:.1f} с | лицо: вступление | – | «{ph[sg['j'][0]][2][:60]}» |")
            continue
        shown.add(i)
        conf = sum(sim[i][j] for j in sg["j"]) / top / len(sg["j"])
        kw = sorted({x for j in sg["j"] for x in hits[i][j]}, key=lambda x: -wt(x))[:6]
        flags = []
        if conf < 0.12:
            flags.append("❓ слов слайда в речи почти нет – проверить по смыслу")
        if b - a > 20:
            flags.append("⏱ дольше 20 с – разбить: view на часть слайда / лицо / карточка")
        if b - a < 2.5:
            flags.append("⚡ меньше 2,5 с – не успеют прочитать")
        draft.append({"from": round(a, 2), "kind": "slide", "src": sl[i]["img"], "n": i + 1, "total": N, "_conf": round(conf, 2)})
        rows.append(f"| {i + 1} | {fmt(a)} | {b - a:.1f} с | {conf:.2f} {'·'.join(kw) or '–'} | {' '.join(flags) or '✅'} "
                    f"| «{ph[sg['j'][0]][2][:60]}» |")
    miss = [s["n"] for k, s in enumerate(sl) if k not in shown]
    L = [f"# Слайды ↔ речь – {p.name}", "",
         f"Урок {fmt(total)}, слайдов {N}, фраз {M}. Начало презентации: {fmt(ph[s0][0]) if M else '–'}"
         f"{' (задано --start)' if start else ' (с начала ролика)'}. В среднем на слайд {(total - (ph[s0][0] if M else 0)) / N:.1f} с.",
         "", "Черновик по словам и порядку. Каждую строку проверить: открыть слайд, прочитать фразы, решить (STYLE.md lesson, раздел 8).",
         "", "| Слайд | С | Длина | Уверенность, общие слова | Флаги | Первая фраза |", "|---|---|---|---|---|---|", *rows, ""]
    if miss:
        L += [f"**Не показаны:** слайды {', '.join(map(str, miss))} – автор их не проговаривает или черновик промахнулся. "
              "Посмотреть каждый: вставить коротко рядом с соседним по смыслу или убрать с причиной.", ""]
    L += ["## Фразы ролика с черновым слайдом", ""]
    L += [f"- {fmt(a)}  [{'лицо' if path[j] is None else path[j] + 1}]  {t}" for j, (a, _, t) in enumerate(ph)]
    (p / "slides_match.md").write_text("\n".join(L), encoding="utf-8")
    save(p, "slides_draft.json", draft)
    print(f"→ {p / 'slides_match.md'}, {p / 'slides_draft.json'}")
    print(f"Сомнительных строк: {sum('❓' in r or '⏱' in r or '⚡' in r for r in rows)}, не показано слайдов: {len(miss)}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("project")
    ap.add_argument("file", nargs="?")
    ap.add_argument("--match", action="store_true")
    ap.add_argument("--start", help="где в ролике начинается презентация: секунды или слово")
    a = ap.parse_args()
    p = project(a.project)
    if a.match:
        st = a.start
        if st is not None:
            try:
                st = float(st)
            except ValueError:
                pass
        match(p, st)
    elif a.file:
        extract(p, a.file)
    else:
        ap.error("дай файл презентации или --match")


if __name__ == "__main__":
    main()
