"""Draw the README's illustrations and charts, in a light and a dark version.

    python docs/make_readme_art.py

Development only, like make_icon.py: it needs Pillow. Writes docs/readme/*.png.
The README shows each pair through <picture>, so GitHub picks the one that
matches the reader's theme.

Every serial, site and model drawn here is invented. The two charts are real:
the read routes come from the 2026-09-30 archive audit (81,273 forms) and the
timings from the 20,000-form benchmark in CLAUDE.md. Update them by hand when
those are measured again.
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).resolve().parent / "readme"
FONTS = Path(r"C:\Windows\Fonts")
S = 2                     # everything is drawn at twice its displayed size

# The app's own palette (theme.CLINICAL), plus the two chart colours, which
# the dataviz validator passes on these surfaces (#008f80 / #26a090).
THEMES = {
    "light": dict(page="#eef3f3", card="#ffffff", card2="#f4f8f8", border="#d2dddd",
                  grid="#e1e9e9", text="#132123", muted="#4a5d61", faint="#6f8286",
                  accent="#0b7a6e", soft="#d9f1ed", on_soft="#0b4d45", mark="#008f80",
                  amber="#b87400", amber_tint="#fcf3df", amber_text="#8f5f00",
                  violet="#6b5bd2", violet_tint="#e9e6fb"),
    "dark": dict(page="#121719", card="#192024", card2="#1f272c", border="#2c373d",
                 grid="#253036", text="#e4ecee", muted="#9cabb1", faint="#7b8b92",
                 accent="#3cc2b0", soft="#1b3a37", on_soft="#bff0e8", mark="#26a090",
                 amber="#e6b04a", amber_tint="#2b2717", amber_text="#e6b04a",
                 violet="#a99cf2", violet_tint="#272440"),
}

# Excel stays Excel in both themes: a register is a light workbook.
SHEET = dict(bg="#ffffff", grid="#d9d9d9", head="#f3f3f3", head_text="#595959",
             text="#1f1f1f", flag="#ffe699", note="#fffbe0", note_border="#a6a6a6",
             triangle="#c00000", title="#1f6f65")


def font(size: float, weight: str = "regular") -> ImageFont.FreeTypeFont:
    files = {"regular": "segoeui.ttf", "semibold": "seguisb.ttf", "bold": "segoeuib.ttf",
             "light": "segoeuil.ttf", "mono": "CascadiaMono.ttf"}
    path = FONTS / files[weight]
    if not path.exists():
        path = FONTS / ("consola.ttf" if weight == "mono" else "segoeui.ttf")
    return ImageFont.truetype(str(path), int(size * S))


def canvas(width: int, height: int) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    image = Image.new("RGBA", (width * S, height * S), (0, 0, 0, 0))
    return image, ImageDraw.Draw(image)


def box(draw, xy, fill=None, outline=None, radius=10, width=1):
    x0, y0, x1, y1 = (v * S for v in xy)
    draw.rounded_rectangle((x0, y0, x1, y1), radius=radius * S, fill=fill,
                           outline=outline, width=width * S)


def text(draw, xy, words, fnt, fill, anchor="la"):
    draw.text((xy[0] * S, xy[1] * S), words, font=fnt, fill=fill, anchor=anchor)


def width_of(words: str, fnt) -> float:
    return fnt.getlength(words) / S


def badge(draw, x, y, number, t, r=11):
    draw.ellipse(((x - r) * S, (y - r) * S, (x + r) * S, (y + r) * S), fill=t["accent"])
    on = "#ffffff" if t is THEMES["light"] else "#04201c"
    text(draw, (x, y + 0.5), str(number), font(12, "bold"), on, "mm")


def save(image: Image.Image, name: str, theme: str) -> None:
    OUT.mkdir(exist_ok=True)
    path = OUT / f"{name}-{theme}.png"
    image.save(path, optimize=True)
    print(path.relative_to(OUT.parent.parent))


def col_index(ref: str) -> tuple[int, int]:
    letters = "".join(c for c in ref if c.isalpha())
    return ord(letters) - ord("A"), int(ref[len(letters):])


# ── 1. Same six facts, a different place on every form ───────────────────────

LAYOUTS = [
    ("Defibrillator", "AC", {"Model": "E15", "S.N": "K15", "Date": "E13", "Status": "G24"}),
    ("ECG", "AF", {"Model": "D32", "S.N": "J32", "Date": "D30", "Status": "F41"}),
    ("Ultrasound", "BB", {"Model": "F17", "S.N": "L17", "Date": "F15", "Status": "H30"}),
    ("Sphygmomanometer", "CE", {"Model": "E47", "S.N": "K47", "Date": "E43", "Status": "H59"}),
]


def layouts(theme: str) -> None:
    t = THEMES[theme]
    W, H = 900, 520
    image, draw = canvas(W, H)
    box(draw, (0, 0, W, H), fill=t["page"], radius=14)
    colours = {"Model": (t["soft"], t["accent"]), "S.N": (t["mark"], t["mark"]),
               "Status": (t["amber_tint"], t["amber"]), "Date": (t["violet_tint"], t["violet"])}
    pad, gap = 20, 14
    card_w = (W - 2 * pad - 3 * gap) / 4
    cols, rows = 12, 60
    for i, (name, code, cells) in enumerate(LAYOUTS):
        x0 = pad + i * (card_w + gap)
        y0 = pad
        box(draw, (x0, y0, x0 + card_w, H - 52), fill=t["card"], outline=t["border"], radius=10)
        text(draw, (x0 + 14, y0 + 14), name, font(13, "semibold"), t["text"])
        chip = font(10, "mono")
        cw = width_of(code, chip) + 12
        box(draw, (x0 + card_w - 14 - cw, y0 + 14, x0 + card_w - 14, y0 + 32),
            fill=t["card2"], outline=t["border"], radius=5)
        text(draw, (x0 + card_w - 14 - cw / 2, y0 + 23), code, chip, t["muted"], "mm")
        gx, gy = x0 + 26, y0 + 58
        gw, gh = card_w - 40, 300
        cell_w, cell_h = gw / cols, gh / rows
        letters = font(8, "mono")
        for c in range(cols):
            text(draw, (gx + (c + 0.5) * cell_w, gy - 9), chr(65 + c), letters, t["faint"], "mm")
        for r in (1, 15, 30, 45, 60):
            text(draw, (gx - 5, gy + (r - 0.5) * cell_h), str(r), letters, t["faint"], "rm")
        for c in range(cols + 1):
            x = gx + c * cell_w
            draw.line((x * S, gy * S, x * S, (gy + gh) * S), fill=t["grid"], width=1)
        for r in range(0, rows + 1, 5):
            y = gy + r * cell_h
            draw.line((gx * S, y * S, (gx + gw) * S, y * S), fill=t["grid"], width=1)
        for field, ref in cells.items():
            c, r = col_index(ref)
            fill, edge = colours[field]
            x, y = gx + c * cell_w, gy + (r - 1) * cell_h
            box(draw, (x - 1, y - 1.5, x + cell_w + 1, y + cell_h + 1.5), fill=fill,
                outline=edge, radius=2, width=1)
        legend_y = gy + gh + 18
        small = font(10.5)
        for j, field in enumerate(("Model", "S.N", "Date", "Status")):
            label = {"S.N": "Serial"}.get(field, field)
            row_y = legend_y + j * 17
            text(draw, (gx - 14, row_y), label, small, t["muted"])
            text(draw, (gx + 52, row_y), cells[field], font(10.5, "mono"), t["text"])
    # Legend for the colours, once, beneath the cards.
    lx, ly = pad + 4, H - 26
    for field, label in (("Model", "Model"), ("S.N", "Serial number"), ("Date", "Date"),
                         ("Status", "Status box")):
        fill, edge = colours[field]
        box(draw, (lx, ly - 6, lx + 22, ly + 6), fill=fill, outline=edge, radius=3)
        text(draw, (lx + 30, ly), label, font(11.5), t["muted"], "lm")
        lx += 30 + width_of(label, font(11.5)) + 26
    text(draw, (W - pad - 4, ly), "columns A–L, rows 1–60 of each form", font(11.5),
         t["faint"], "rm")
    save(image, "layouts", theme)


# ── 2. One form becomes two register rows ────────────────────────────────────

def form_to_row(theme: str) -> None:
    t = THEMES[theme]
    W, H = 900, 540
    image, draw = canvas(W, H)
    box(draw, (0, 0, W, H), fill=t["page"], radius=14)

    # The form, as Excel draws it: a light sheet inside the theme card.
    sx, sy, sw = 24, 46, W - 48
    text(draw, (sx, 24), "W01-AGH011-0326.xlsx", font(13, "semibold"), t["text"], "lm")
    text(draw, (sx + width_of("W01-AGH011-0326.xlsx", font(13, "semibold")) + 10, 24),
         "a patient monitor's inspection form", font(12), t["muted"], "lm")
    cols = "ABCDEFGHIJKL"
    head_w, row_h = 30, 22
    cell_w = (sw - head_w) / len(cols)
    shown = [14, 15, 16, 17, 18, 19, 20, 21, None, 37, 38, 39, 40]
    sh = row_h * (len(shown) + 1)
    box(draw, (sx, sy, sx + sw, sy + sh), fill=SHEET["bg"], outline=SHEET["grid"], radius=4)
    draw.rectangle((sx * S, sy * S, (sx + sw) * S, (sy + row_h) * S), fill=SHEET["head"])
    tiny = font(10)
    for i, letter in enumerate(cols):
        x = sx + head_w + i * cell_w
        text(draw, (x + cell_w / 2, sy + row_h / 2), letter, tiny, SHEET["head_text"], "mm")
        draw.line((x * S, sy * S, x * S, (sy + sh) * S), fill=SHEET["grid"], width=1)
    ys = {}
    for k, r in enumerate(shown, start=1):
        y = sy + k * row_h
        draw.line((sx * S, y * S, (sx + sw) * S, y * S), fill=SHEET["grid"], width=1)
        label = "…" if r is None else str(r)
        text(draw, (sx + head_w / 2, y + row_h / 2), label, tiny, SHEET["head_text"], "mm")
        if r is not None:
            ys[r] = y

    def cell(ref, span=1):
        c, r = col_index(ref)
        return sx + head_w + c * cell_w, ys[r], sx + head_w + (c + span) * cell_w, ys[r] + row_h

    caption = font(10.5)
    for ref, words in {"A16": "Date of receipt:", "A18": "Model:", "A20": "Manufacturer:",
                       "H18": "Serial No.:", "H20": "Location:", "B38": "ECG status",
                       "H38": "NIBP status"}.items():
        x0, y0, _, _ = cell(ref)
        text(draw, (x0 + 5, y0 + row_h / 2), words, caption, "#404040", "lm")
    value = font(11, "semibold")
    filled = [("E16", 2, "12-03-2026", 5), ("E18", 3, "IntelliVue MX450", 2),
              ("E20", 2, "Philips", 1), ("K18", 2, "PM-88213", 3), ("K20", 2, "ICU", 4),
              ("D39", 2, "Pass", 6), ("J39", 2, "Pass", 7)]
    for ref, span, words, number in filled:
        x0, y0, x1, y1 = cell(ref, span)
        draw.rectangle((x0 * S + 2, y0 * S + 2, x1 * S - 1, y1 * S - 1), fill=t["soft"]
                       if theme == "light" else "#d9f1ed", outline="#0b7a6e", width=2)
        text(draw, (x0 + 6, (y0 + y1) / 2), words, value, "#0b4d45", "lm")
        badge(draw, x1 - 2, y0, number, THEMES["light"], r=9)

    # The two rows it becomes.
    ry = sy + sh + 46
    text(draw, (sx, ry - 18), "device list.xlsx", font(13, "semibold"), t["text"], "lm")
    text(draw, (sx + width_of("device list.xlsx", font(13, "semibold")) + 10, ry - 18),
         "two rows: the monitor, and its NIBP module on the same serial",
         font(12), t["muted"], "lm")
    headers = [("No.", 36, None), ("Device", 128, None), ("Manufacturer", 90, 1),
               ("Model", 120, 2), ("S.N", 82, 3), ("Location", 66, 4),
               ("Code", 132, None), ("Date", 84, 5), ("Status", 0, 6)]
    fixed = sum(w for _, w, _ in headers)
    headers[-1] = ("Status", sw - fixed, 6)
    rows = [["1", "Patient Monitor", "Philips", "IntelliVue MX450", "PM-88213", "ICU",
             "W01-AGH011-0326", "12-03-2026", "Pass"],
            ["2", "NIBP Module", "Philips", "IntelliVue MX450", "PM-88213", "ICU",
             "W01-AGCB011-0326", "12-03-2026", "Pass"]]
    rh = 28
    box(draw, (sx, ry, sx + sw, ry + rh * 3), fill=SHEET["bg"], outline=SHEET["grid"], radius=4)
    draw.rectangle((sx * S + 1, ry * S + 1, (sx + sw) * S - 1, (ry + rh) * S),
                   fill=SHEET["title"])
    x = sx
    for i, (name, w, number) in enumerate(headers):
        text(draw, (x + 7, ry + rh / 2), name, font(11, "semibold"), "#ffffff", "lm")
        if number:
            badge(draw, x + w - 13, ry, number, THEMES["light"], r=9)
        for k, row in enumerate(rows, start=1):
            y = ry + k * rh
            changed = k == 2 and name in ("Device", "Code")
            if changed or (k == 2 and name == "Status"):
                draw.rectangle(((x + 1) * S, (y + 1) * S, (x + w - 1) * S, (y + rh - 1) * S),
                               fill="#eef6f5")
            words = row[i]
            fnt = font(10.5, "semibold" if changed else "regular")
            while width_of(words, fnt) > w - 12 and len(words) > 3:
                words = words[:-2] + "…"
            text(draw, (x + 7, y + rh / 2), words, fnt, SHEET["text"], "lm")
        if i:
            draw.line((x * S, (ry + rh) * S, x * S, (ry + 3 * rh) * S), fill=SHEET["grid"])
        x += w
    draw.line((sx * S, (ry + 2 * rh) * S, (sx + sw) * S, (ry + 2 * rh) * S), fill=SHEET["grid"])
    tail = "NIBP status: the module row's own verdict"
    text(draw, (sx + sw, ry + 3 * rh + 22), tail, font(11), t["muted"], "rm")
    badge(draw, sx + sw - width_of(tail, font(11)) - 16, ry + 3 * rh + 22, 7, t, r=9)
    text(draw, (sx, ry + 3 * rh + 22), "Numbers match each value to its box on the form. "
         "Invented data.", font(11), t["faint"], "lm")
    save(image, "form-to-row", theme)


# ── 3. The register marks what it doubts ─────────────────────────────────────

def notes(theme: str) -> None:
    t = THEMES[theme]
    W, H = 900, 292
    image, draw = canvas(W, H)
    box(draw, (0, 0, W, H), fill=t["page"], radius=14)
    sx, sy, sw = 24, 24, W - 48
    headers = [("No.", 36), ("Device", 104), ("Manufacturer", 96), ("Model", 112),
               ("S.N", 92), ("Location", 70), ("Code", 136), ("Date", 84), ("Status", 0)]
    headers[-1] = ("Status", sw - sum(w for _, w in headers))
    rows = [
        ["1", "Centrifuge", "Hettich", "EBA 200", "CF-30418", "Lab", "W01-AS002-0326",
         "11-03-2026", "JTE-40"],
        ["2", "ECG", "Schiller", "AT-102 G2", "0", "ER", "W01-AF007-0326",
         "12-03-2026", "Pass"],
        ["3", "Syringe", "B. Braun", "Perfusor Space", "SY-40177", "NICU", "W01-BZ014-0326",
         "12-03-2026", "Pass"],
        ["4", "Defibrillator", "Zoll", "R Series", "DF-20481", "ICU", "W01-AC004-0326",
         "13-03-2026", "Pass"],
    ]
    flagged = {(0, "Status"), (1, "Manufacturer"), (1, "Model"), (1, "S.N"),
               (1, "Location"), (1, "Date"), (2, "Code")}
    rh = 30
    sheet_h = rh * (len(rows) + 1)
    box(draw, (sx, sy, sx + sw, sy + sheet_h), fill=SHEET["bg"], outline=SHEET["grid"], radius=4)
    draw.rectangle((sx * S + 1, sy * S + 1, (sx + sw) * S - 1, (sy + rh) * S), fill=SHEET["title"])
    x = sx
    spots = {}
    for i, (name, w) in enumerate(headers):
        text(draw, (x + 7, sy + rh / 2), name, font(11, "semibold"), "#ffffff", "lm")
        for k, row in enumerate(rows):
            y = sy + (k + 1) * rh
            if (k, name) in flagged:
                draw.rectangle(((x + 1) * S, (y + 1) * S, (x + w - 1) * S, (y + rh - 1) * S),
                               fill=SHEET["flag"])
                tri = [((x + w - 1) * S, (y + 1) * S), ((x + w - 9) * S, (y + 1) * S),
                       ((x + w - 1) * S, (y + 9) * S)]
                draw.polygon(tri, fill=SHEET["triangle"])
                spots[(k, name)] = (x, y, w)
            text(draw, (x + 7, y + rh / 2), row[i], font(10.5), SHEET["text"], "lm")
        if i:
            draw.line((x * S, (sy + rh) * S, x * S, (sy + sheet_h) * S), fill=SHEET["grid"])
        x += w
    for k in range(2, len(rows) + 1):
        y = sy + k * rh
        draw.line((sx * S, y * S, (sx + sw) * S, y * S), fill=SHEET["grid"])

    # The hovered note, where Excel puts it: up and to the right of its cell.
    x, y, w = spots[(1, "S.N")]
    lines = ["Calist:", "The Serial No. box is empty on the 'data entry'",
             "tab, so this was read from the 'Cover Report' tab."]
    nx, ny, nw = x + w + 14, y + 10, 300
    nh = 14 + 17 * len(lines)
    draw.line(((x + w - 2) * S, (y + 3) * S, nx * S, (ny + 6) * S), fill="#7f7f7f", width=S)
    draw.rectangle((nx * S + 3 * S, ny * S + 3 * S, (nx + nw) * S + 3 * S, (ny + nh) * S + 3 * S),
                   fill="#00000022")
    draw.rectangle((nx * S, ny * S, (nx + nw) * S, (ny + nh) * S), fill=SHEET["note"],
                   outline=SHEET["note_border"], width=S)
    for j, line in enumerate(lines):
        text(draw, (nx + 8, ny + 8 + 17 * j), line, font(10.5, "semibold" if j == 0 else "regular"),
             "#1f1f1f")

    key = [("Row 1", "JTE-40 is an engineer's code from the \"Tested by\" box, not a verdict."),
           ("Row 2", "the serial was left empty on the data tab; the details came from the "
                     "Cover Report."),
           ("Row 3", "\" (2)\" after the date was left out; the note names the real file.")]
    ky = sy + sheet_h + 26
    for j, (label, words) in enumerate(key):
        box(draw, (sx, ky + j * 24 - 7, sx + 18, ky + j * 24 + 7), fill=SHEET["flag"],
            outline="#d9b44a", radius=3)
        text(draw, (sx + 28, ky + j * 24), label, font(11.5, "semibold"), t["text"], "lm")
        text(draw, (sx + 76, ky + j * 24), words, font(11.5), t["muted"], "lm")
    text(draw, (sx, H - 18), "Amber means \"kept, but look at this\". Hover the cell in Excel "
         "to read why. Invented data.", font(11), t["faint"], "lm")
    save(image, "notes", theme)


# ── 4. Anatomy of a file name ────────────────────────────────────────────────

def filename(theme: str) -> None:
    t = THEMES[theme]
    W, H = 900, 200
    image, draw = canvas(W, H)
    box(draw, (0, 0, W, H), fill=t["page"], radius=14)
    big = font(46, "mono")
    parts = [("G302", t["violet"], "site code", "a letter, then digits"),
             ("-", None, None, None),
             ("AGH", t["accent"], "device", "Patient Monitor"),
             ("001", t["mark"], "unit", "number 1"),
             ("-", None, None, None),
             ("0425", t["amber"], "round", "April 2025, as MMYY"),
             (".xlsx", None, None, None)]
    total = sum(width_of(p[0], big) for p in parts)
    x = (W - total) / 2
    y = 78
    for words, colour, label, hint in parts:
        w = width_of(words, big)
        text(draw, (x, y), words, big, colour or t["faint"], "ls")
        if label:
            draw.line(((x + 2) * S, (y + 14) * S, (x + w - 2) * S, (y + 14) * S),
                      fill=colour, width=3 * S)
            cx = x + w / 2
            text(draw, (cx, y + 40), label, font(13, "semibold"), t["text"], "mm")
            text(draw, (cx, y + 60), hint, font(11.5), t["muted"], "mm")
        x += w
    text(draw, (W / 2, H - 26), "The device code alone picks the form's layout: "
         "everything after the first dash, letters only.", font(12), t["faint"], "mm")
    save(image, "filename", theme)


# ── 5 and 6. Charts ──────────────────────────────────────────────────────────

def bar_chart(theme: str, name: str, title: str, subtitle: str,
              bars: list[tuple[str, float, str]], unit_note: str) -> None:
    """Horizontal bars, one series: the title names it, so no legend. Values
    at the tips; text in text colours, never the bar colour."""
    t = THEMES[theme]
    W = 900
    bar_h, band = 14, 34
    top = 78
    H = top + band * len(bars) + 46
    image, draw = canvas(W, H)
    box(draw, (0, 0, W, H), fill=t["card"], outline=t["border"], radius=14)
    text(draw, (24, 22), title, font(15, "semibold"), t["text"])
    text(draw, (24, 46), subtitle, font(11.5), t["muted"])
    label_w = max(width_of(b[0], font(12)) for b in bars) + 36
    x0, x1 = 24 + label_w, W - 110
    peak = max(b[1] for b in bars)
    draw.line((x0 * S, (top - 6) * S, x0 * S, (top + band * len(bars) - 10) * S),
              fill=t["grid"], width=S)
    for i, (label, val, shown) in enumerate(bars):
        cy = top + i * band + band / 2 - 8
        text(draw, (x0 - 12, cy), label, font(12), t["text"], "rm")
        length = max(2.0, (x1 - x0) * val / peak)
        bx0, by0, bx1, by1 = x0 * S, (cy - bar_h / 2) * S, (x0 + length) * S, (cy + bar_h / 2) * S
        r = min(4 * S, int((bx1 - bx0) / 2))
        draw.rounded_rectangle((bx0, by0, bx1, by1), radius=r, fill=t["mark"])
        if bx1 - bx0 > r:
            draw.rectangle((bx0, by0, bx0 + r, by1), fill=t["mark"])   # square at the base
        text(draw, (x0 + length + 8, cy), shown, font(12, "semibold"), t["text"], "lm")
    text(draw, (24, H - 22), unit_note, font(11), t["faint"], "lm")
    save(image, name, theme)


def charts(theme: str) -> None:
    routes = [("Its own cell map", 74554), ("An older layout we wrote down", 4600),
              ("Printed labels, on another tab", 882), ("Not a form (refused, listed)", 444),
              ("Word certificate beside it", 401), ("Printed labels, same tab", 240),
              ("Nothing fits (kept, marked amber)", 152)]
    total = sum(n for _, n in routes)
    bar_chart(theme, "chart-routes", "How 81,273 archive forms were read",
              "The first route that gives a believable record wins. Four years of rounds, "
              "2023 to 2026.",
              [(label, n, f"{n:,}  ·  {100 * n / total:.1f}%") for label, n in routes],
              "From tools/audit.py, run over the whole archive on 30 September 2026.")
    bar_chart(theme, "chart-time", "Where the time goes on 20,000 forms",
              "One run, 38.3 seconds end to end, on an ordinary office desktop.",
              [("Reading the forms", 28.0, "28.0 s"), ("Writing the register", 9.6, "9.6 s"),
               ("Sorting", 0.5, "0.5 s"), ("Removing duplicate serials", 0.2, "0.2 s")],
              "Writing is inside openpyxl's save, the one part left that Calist cannot "
              "make faster.")


if __name__ == "__main__":
    for theme in THEMES:
        layouts(theme)
        form_to_row(theme)
        notes(theme)
        filename(theme)
        charts(theme)
