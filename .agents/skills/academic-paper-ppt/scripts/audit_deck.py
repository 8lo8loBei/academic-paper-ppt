#!/usr/bin/env python3
"""Audit a PPTX against the academic-paper-ppt layout rules.

Usage: python audit_deck.py deck.pptx [--json]
Needs python-pptx (and Pillow for picture ink-bbox). Exit code 1 if any FAIL.
Machine-checkable rules only; visual judgement still needs the rendered pages.
"""
import io
import re
import sys
import json
import zipfile

from pptx import Presentation
from pptx.util import Emu

try:
    from PIL import Image, ImageChops
except ImportError:  # ink-bbox check is skipped without Pillow
    Image = None

IN = 914400
MIN_BODY_PT, MIN_ANY_PT = 17.0, 12.0
CONTENT_TOP_MAX, FIG_BOTTOM_MIN, TEXT_BOTTOM_MIN = 2.10, 5.90, 5.20
SPAN_MIN, FIG_AREA_MIN, COVER_FIG_MIN, INK_MIN = 9.5, 0.25, 0.15, 0.85
HEADER_END, FOOTER_START = 1.75, 6.55
PAGE_NO = re.compile(r"^\s*(\d{1,3}\s*/\s*\d{1,3}|\d{1,3})\s*$")
META_TITLE = re.compile(r"(的读法|解读$|结果分析|^小结$|Thank\s*you|谢谢)", re.I)
CAPTION = re.compile(r"^\s*(Fig|Figure|Table|Tab|Eq|Extended Data|Supplementary|\[\d+\])", re.I)
LABEL = re.compile(r"^[^，。,.；;]{1,8}[：:]\s*\S")


def leaves(shapes, dx=0, dy=0):
    """Yield (shape, left, top, width, height) in inches, flattening groups."""
    for sh in shapes:
        if sh.shape_type == 6:  # group
            yield sh, sh.left / IN, sh.top / IN, sh.width / IN, sh.height / IN
            continue
        if sh.width is None or sh.left is None:
            continue
        yield sh, sh.left / IN, sh.top / IN, sh.width / IN, sh.height / IN


def ink_ratio(blob):
    """Share of the picture's own area covered by non-white ink bbox."""
    if Image is None:
        return None
    try:
        im = Image.open(io.BytesIO(blob)).convert("RGB")
    except Exception:
        return None
    bg = Image.new("RGB", im.size, (255, 255, 255))
    diff = ImageChops.difference(im, bg).convert("L").point(lambda v: 255 if v > 7 else 0)
    box = diff.getbbox()
    if not box:
        return 0.0
    return ((box[2] - box[0]) * (box[3] - box[1])) / float(im.size[0] * im.size[1])


def audit(path):
    prs = Presentation(path)
    W, H = prs.slide_width / IN, prs.slide_height / IN
    n = len(prs.slides)
    out = []

    def add(slide, level, rule, msg):
        out.append({"slide": slide, "level": level, "rule": rule, "msg": msg})

    if abs(W / H - 16 / 9) > 0.01:
        add(0, "FAIL", "canvas", f"画布 {W:.2f}×{H:.2f} 英寸，不是 16:9")

    with zipfile.ZipFile(path) as z:
        xml = {m: z.read(m).decode("utf8", "ignore") for m in z.namelist()
               if re.match(r"ppt/slides/slide\d+\.xml$", m)}
    math_total = sum(x.count("<m:oMath>") + x.count("<m:oMath ") for x in xml.values())

    for i, s in enumerate(prs.slides, 1):
        items = list(leaves(s.shapes))
        # --- notes
        notes = s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else ""
        if not notes:
            add(i, "FAIL", "notes", "没有演讲者备注")
        elif len(notes) < 30:
            add(i, "WARN", "notes", f"备注只有 {len(notes)} 字")
        # --- text checks
        body_lines = 0
        for sh, l, t, w, h in items:
            if not getattr(sh, "has_text_frame", False) or not sh.has_text_frame:
                continue
            text = sh.text_frame.text.strip()
            if not text:
                continue
            if PAGE_NO.match(text) and t > FOOTER_START - 0.2:
                add(i, "FAIL", "page-number", f"页脚出现页码/计数「{text}」")
            if t < HEADER_END and META_TITLE.search(text):
                add(i, "FAIL", "meta-title", f"元标题「{text[:20]}」")
            if i == n and META_TITLE.search(text) and len(text) < 12:
                add(i, "FAIL", "last-page", f"末页是「{text}」，应为总结或参考文献")
            in_body = HEADER_END <= t < FOOTER_START
            for p in sh.text_frame.paragraphs:
                ptxt = "".join(r.text for r in p.runs).strip()
                if not ptxt:
                    continue
                sizes = [r.font.size.pt for r in p.runs if r.font.size]
                if in_body and i > 1 and len(ptxt) >= 14:
                    body_lines += 1
                    if LABEL.match(ptxt):
                        add(i, "WARN", "bullet-label", f"分点疑似带前置标签「{ptxt[:14]}…」")
                    if sizes and min(sizes) < MIN_ANY_PT - 0.01:
                        add(i, "FAIL", "font-size", f"{min(sizes):g} pt < {MIN_ANY_PT:g}：「{ptxt[:14]}…」")
                    elif sizes and min(sizes) < MIN_BODY_PT and not CAPTION.match(ptxt):
                        add(i, "WARN", "font-size", f"{min(sizes):g} pt；若是正文分点须 ≥{MIN_BODY_PT:g}：「{ptxt[:14]}…」")
                elif sizes and min(sizes) < MIN_ANY_PT - 0.01:
                    add(i, "FAIL", "font-size", f"{min(sizes):g} pt < {MIN_ANY_PT:g}：「{ptxt[:14]}」")
                for r in p.runs:
                    try:
                        c = r.font.color
                        if c is not None and c.type == 1:
                            rgb = c.rgb
                            if min(rgb[0], rgb[1], rgb[2]) > 0x8C and max(rgb) - min(rgb) < 30 \
                                    and min(rgb[0], rgb[1], rgb[2]) < 0xF0:
                                add(i, "WARN", "light-text", f"浅灰文字 #{rgb}：「{r.text[:12]}」")
                                break
                    except Exception:
                        pass
        if body_lines > 8:
            add(i, "WARN", "density", f"正文段落 {body_lines} 段，可能超过 8 行")
        # --- figures
        figs = [(sh, l, t, w, h) for sh, l, t, w, h in items
                if sh.shape_type in (13, 6) and w * h / (W * H) > 0.03 and t < FOOTER_START]
        fig_area = sum(w * h for _, _, _, w, h in figs) / (W * H)
        for sh, l, t, w, h in figs:
            if sh.shape_type == 13:
                r = ink_ratio(sh.image.blob)
                if r is not None and r < INK_MIN:
                    add(i, "FAIL", "ink-crop", f"图片墨迹只占图框 {r:.0%}，先紧裁白边再放置")
        # --- geometry of the content band (skip cover)
        if i == 1:
            if figs and fig_area < COVER_FIG_MIN:
                add(i, "WARN", "cover-figure", f"封面主图占页 {fig_area:.0%} < {COVER_FIG_MIN:.0%}")
            tops = [t for _, _, t, _, h in items if h > 0.05]
            bots = [t + h for _, _, t, _, h in items]
            if tops and min(tops) > 0.8:
                add(i, "FAIL", "cover-top", f"封面顶带空：最上对象在 {min(tops):.2f} 英寸")
            if bots and max(bots) < 5.9:
                add(i, "FAIL", "cover-bottom", f"封面底带空：最下对象止于 {max(bots):.2f} 英寸")
            continue
        band = [(l, t, w, h) for _, l, t, w, h in items
                if t + h / 2 > HEADER_END and t + h / 2 < FOOTER_START and h > 0.05]
        if not band:
            add(i, "WARN", "empty", "内容带没有对象")
            continue
        top = min(t for _, t, _, _ in band)
        bottom = max(t + h for _, t, _, h in band)
        span = max(l + w for l, _, w, _ in band) - min(l for l, _, _, _ in band)
        if top > CONTENT_TOP_MAX + 0.15:
            add(i, "WARN", "content-top", f"内容上沿 {top:.2f} > {CONTENT_TOP_MAX}")
        need = FIG_BOTTOM_MIN if figs else TEXT_BOTTOM_MIN
        if bottom < need - 0.05:
            add(i, "FAIL", "whitespace", f"内容下沿 {bottom:.2f} < {need}，页面下方留白过多")
        if span < SPAN_MIN:
            add(i, "FAIL", "whitespace", f"内容横向只占 {span:.1f} 英寸 < {SPAN_MIN}")
        if figs and fig_area < FIG_AREA_MIN:
            add(i, "FAIL", "figure-area", f"图占页面积 {fig_area:.0%} < {FIG_AREA_MIN:.0%}")
    return {"file": path, "slides": n, "omml_formulas": math_total, "findings": out}


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        return 0
    rep = audit(sys.argv[1])
    if "--json" in sys.argv:
        print(json.dumps(rep, ensure_ascii=False, indent=2))
    else:
        print(f"{rep['file']}: {rep['slides']} 页，OMML 公式 {rep['omml_formulas']} 个")
        for f in rep["findings"]:
            print(f"  [{f['level']}] p{f['slide']:>2} {f['rule']}: {f['msg']}")
        fails = sum(f["level"] == "FAIL" for f in rep["findings"])
        warns = len(rep["findings"]) - fails
        print(f"FAIL {fails} / WARN {warns}")
    return 1 if any(f["level"] == "FAIL" for f in rep["findings"]) else 0


if __name__ == "__main__":
    sys.exit(main())
