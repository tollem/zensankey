#!/usr/bin/env python3
"""Renders the layouts printed by tests/SkTest.bas (via build/test_layout.py) as SVG files,
to preview the appearance of the add-in without Office.

Usage: python3 build/test_layout.py > layout.txt && python3 build/preview_svg.py layout.txt OUT_DIR
"""
import sys, os, html


def parse(text):
    sections, cur = [], None
    for line in text.splitlines():
        if line.startswith("### "):
            cur = {"name": line[4:].strip(), "cols": {}, "nodes": [], "links": [], "box": None, "opt": None}
            sections.append(cur)
        elif cur is not None and "|" in line:
            p = line.split("|")
            if p[0] == "BOX":
                cur["box"] = [float(x) for x in p[1:]]
            elif p[0] == "C":
                cur["cols"][int(p[1])] = float(p[2])
            elif p[0] == "N":
                cur["nodes"].append(dict(label=p[1], col=int(p[2]), x=float(p[3]), y=float(p[4]), h=float(p[5]),
                                         amount=p[6], text=p[7], fill=p[8], lx=float(p[9]), ly=float(p[10]),
                                         lw=float(p[11]), lh=float(p[12]), side=p[13]))
            elif p[0] == "L":
                cur["links"].append(dict(src=p[1], dst=p[2], y0=float(p[3]), y1=float(p[4]), h=float(p[5]),
                                         c1=p[6], c2=p[7], amount=p[8]))
            elif p[0] == "OPT":
                cur["opt"] = p[1:]
    return [s for s in sections if s["box"]]


def wrap(label, max_chars):
    words, lines, cur = label.split(), [], ""
    for w in words:
        if cur and len(cur) + 1 + len(w) > max_chars:
            lines.append(cur); cur = w
        else:
            cur = (cur + " " + w).strip()
    if cur:
        lines.append(cur)
    return lines


def render(sec):
    boxW, boxH, areaL, areaT, areaW, areaH, fs, pad = sec["box"]
    opt = sec["opt"]
    palette, color_by, link_color, shape, opacity, labels, fmt, col_titles, title, subtitle, node_w, curv, font, cts = opt[:14]
    node_w = float(node_w); curv = float(curv); opacity = float(opacity); cts = float(cts)
    nodes = {n["label"]: n for n in sec["nodes"]}
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{boxW*2:.0f}" height="{boxH*2:.0f}" viewBox="0 0 {boxW:.1f} {boxH:.1f}" font-family="Arial, Helvetica, sans-serif">',
           f'<rect width="{boxW:.1f}" height="{boxH:.1f}" fill="white"/>']
    for l in sec["links"]:
        s, d = nodes[l["src"]], nodes[l["dst"]]
        x0, x1 = s["x"] + node_w, d["x"]
        y0, y1, h = l["y0"], l["y1"], l["h"]
        if shape == "straight":
            path = f"M{x0:.1f},{y0:.1f} L{x1:.1f},{y1:.1f} L{x1:.1f},{y1+h:.1f} L{x0:.1f},{y0+h:.1f} Z"
        else:
            cx0 = x0 + (x1 - x0) * curv; cx1 = x1 - (x1 - x0) * curv
            path = (f"M{x0:.1f},{y0:.1f} C{cx0:.1f},{y0:.1f} {cx1:.1f},{y1:.1f} {x1:.1f},{y1:.1f} L{x1:.1f},{y1+h:.1f} "
                    f"C{cx1:.1f},{y1+h:.1f} {cx0:.1f},{y0+h:.1f} {x0:.1f},{y0+h:.1f} Z")
        if l["c1"] != l["c2"]:
            gid = f"g{abs(hash(l['src']+l['dst']))}"
            out.append(f'<defs><linearGradient id="{gid}" x1="0" x2="1" y1="0" y2="0"><stop offset="0" stop-color="{l["c1"]}"/><stop offset="1" stop-color="{l["c2"]}"/></linearGradient></defs>')
            fill = f"url(#{gid})"
        else:
            fill = l["c1"]
        out.append(f'<path d="{path}" fill="{fill}" fill-opacity="{opacity}"/>')
    for n in sec["nodes"]:
        out.append(f'<rect x="{n["x"]:.1f}" y="{n["y"]:.1f}" width="{node_w}" height="{n["h"]:.1f}" fill="{n["fill"]}"/>')
    if labels != "none":
        for n in sec["nodes"]:
            anchor = "end" if n["side"] == "L" else "start"
            tx = n["lx"] + n["lw"] - 2 if n["side"] == "L" else n["lx"] + 2
            lines = wrap(n["label"], max(8, int(13 * 1.9))) if labels in ("name_value", "name") else []
            if labels == "name_value":
                lines = lines + [("VALUE", n["text"])]
            elif labels == "value":
                lines = [("VALUE", n["text"])]
            lh = fs * 1.25
            ty = n["ly"] + 1 + fs
            for ln in lines:
                if isinstance(ln, tuple):
                    out.append(f'<text x="{tx:.1f}" y="{ty:.1f}" font-size="{fs}" font-weight="bold" fill="{n["fill"]}" text-anchor="{anchor}">{html.escape(ln[1])}</text>')
                else:
                    out.append(f'<text x="{tx:.1f}" y="{ty:.1f}" font-size="{fs}" fill="#1C2B2D" text-anchor="{anchor}">{html.escape(ln)}</text>')
                ty += lh
    band_title = 0
    if title:
        band_title += fs * 1.6 * 1.5
        out.append(f'<text x="8" y="{8 + fs*1.6:.1f}" font-size="{fs*1.6:.1f}" font-weight="bold" fill="#1C2B2D">{html.escape(title)}</text>')
    if subtitle:
        out.append(f'<text x="8" y="{8 + band_title + fs*1.1:.1f}" font-size="{fs*1.1:.1f}" font-style="italic" fill="#6E787C">{html.escape(subtitle)}</text>')
        band_title += fs * 1.1 * 1.4
    if band_title:
        band_title += 4
    if col_titles:
        for c, t in enumerate(col_titles.split(";")):
            if c in sec["cols"] and t.strip():
                cx = sec["cols"][c] + node_w / 2
                out.append(f'<text x="{cx:.1f}" y="{8 + band_title + cts:.1f}" font-size="{cts}" font-weight="bold" fill="#6E787C" text-anchor="middle">{html.escape(t.strip())}</text>')
    out.append("</svg>")
    return "\n".join(out)


def main():
    text = open(sys.argv[1], encoding="utf-8").read()
    out_dir = sys.argv[2]
    os.makedirs(out_dir, exist_ok=True)
    for sec in parse(text):
        fn = os.path.join(out_dir, sec["name"].replace(" ", "_").replace("(", "").replace(")", "") + ".svg")
        open(fn, "w", encoding="utf-8").write(render(sec))
        print("wrote", fn)


if __name__ == "__main__":
    main()
