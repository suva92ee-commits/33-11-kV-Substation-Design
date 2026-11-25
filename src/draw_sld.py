"""Draw the single line diagram (SLD) of the 33/11 kV substation.

One drawing definition is rendered to three formats:
    drawings/SLD_33_11kV_substation.dxf   (open in AutoCAD / any CAD tool and edit)
    drawings/SLD_33_11kV_substation.png
    drawings/SLD_33_11kV_substation.pdf

Equipment ratings printed on the drawing come from calcs.py (the same numbers the
Excel workbook calculates). The drawing shows the DEFAULT configuration:
two 33 kV incomers, 2 x 10 MVA transformers, sectionalised 11 kV bus, 6 feeders.
Changing the number of transformers or feeders in the inputs requires editing this file.
"""
import datetime
import math
from pathlib import Path

import ezdxf
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Polygon
from ezdxf.enums import TextEntityAlignment

import calcs
import design_config as cfg

W, H = 420.0, 297.0  # A3 landscape, mm

LAYERS = {  # name: (dxf colour, line width mm, dashed)
    "BUS": (7, 0.9, False),
    "LINE": (7, 0.35, False),
    "EQUIP": (7, 0.35, False),
    "DASH": (8, 0.25, True),
    "FRAME": (7, 0.7, False),
    "TEXT": (7, 0.25, False),
}


class Canvas:
    def __init__(self):
        self.items = []

    def line(self, x1, y1, x2, y2, layer="LINE"):
        self.items.append(("line", layer, (x1, y1, x2, y2)))

    def poly(self, pts, layer="EQUIP", closed=True, fill=False):
        self.items.append(("poly", layer, (list(pts), closed, fill)))

    def circle(self, x, y, r, layer="EQUIP"):
        self.items.append(("circle", layer, (x, y, r)))

    def rect(self, x, y, w, h, layer="EQUIP", center=True):
        if center:
            x, y = x - w / 2, y - h / 2
        self.poly([(x, y), (x + w, y), (x + w, y + h), (x, y + h)], layer, True)

    def text(self, x, y, s, h=2.2, rot=0, ha="left", bold=False):
        self.items.append(("text", "TEXT", (x, y, s, h, rot, ha, bold)))


# ---------------------------------------------------------------- symbols
EXT = {"breaker": 2.5, "isolator": 3.0, "ct": 0.0, "fuse": 2.6}


def sym(c, kind, ori, x, y):
    def L(u, w):  # u along the conductor, w across it
        return (x + w, y - u) if ori == "v" else (x + u, y + w)

    if kind == "breaker":
        c.poly([L(-2.5, -2.5), L(-2.5, 2.5), L(2.5, 2.5), L(2.5, -2.5)])
        c.line(*L(-2.5, -2.5), *L(2.5, 2.5), layer="EQUIP")
        c.line(*L(-2.5, 2.5), *L(2.5, -2.5), layer="EQUIP")
    elif kind == "isolator":
        c.line(*L(-3, 0), *L(2.8, 2.4), layer="EQUIP")
        c.line(*L(3, -1.6), *L(3, 1.6), layer="EQUIP")
    elif kind == "ct":
        c.circle(x, y, 2.2)
    elif kind == "fuse":
        c.poly([L(-2.6, -1.2), L(-2.6, 1.2), L(2.6, 1.2), L(2.6, -1.2)])
        c.line(*L(-2.6, 0), *L(2.6, 0), layer="EQUIP")


def chain(c, ori, fixed, a, b, elems, layer="LINE"):
    """Conductor from a to b with symbols. 'v': x=fixed, a>b (top to bottom). 'h': y=fixed, a<b."""
    def P(t):
        return (fixed, a - t) if ori == "v" else (a + t, fixed)

    total, cur = abs(b - a), 0.0
    for pos, kind in sorted((abs(p - a), k) for p, k in elems):
        ext = EXT[kind]
        if pos - ext > cur:
            c.line(*P(cur), *P(pos - ext), layer=layer)
        sym(c, kind, ori, *P(pos))
        cur = pos + ext
    if total > cur:
        c.line(*P(cur), *P(total), layer=layer)


def earth(c, x, y):
    c.line(x, y, x, y - 2)
    for dy, hw in ((2, 3.0), (3.5, 2.0), (5, 1.0)):
        c.line(x - hw, y - dy, x + hw, y - dy)


def arrow_down(c, x, y):
    c.poly([(x - 1.8, y + 3.6), (x + 1.8, y + 3.6), (x, y)], "EQUIP", True, True)


def surge_arrester(c, x, y, d):
    """Tap at (x,y), branch to the side d (+1 right, -1 left), arrester body, earth."""
    bx = x + d * 9
    c.line(x, y, bx, y)
    c.line(bx, y, bx, y - 2)
    c.rect(bx, y - 5.5, 3, 7)
    c.line(bx, y - 2.8, bx, y - 8.2, layer="EQUIP")
    c.line(bx, y - 9, bx, y - 11)
    earth(c, bx, y - 11)
    c.text(bx + (2.4 if d > 0 else -2.4), y - 6.5, "LA", 1.9, ha="left" if d > 0 else "right")


def power_tx(c, x, cy):
    """Two-winding transformer: delta (top) / star (bottom). Returns (top_y, bottom_y)."""
    c.circle(x, cy + 4, 6.5)
    c.circle(x, cy - 4, 6.5)
    t = cy + 4
    c.poly([(x, t + 3), (x - 2.6, t - 1.5), (x + 2.6, t - 1.5)], "EQUIP", True)
    s = cy - 4
    for dx, dy in ((0, 3), (-2.6, -1.5), (2.6, -1.5)):
        c.line(x, s, x + dx, s + dy, layer="EQUIP")
    return cy + 10.5, cy - 10.5


def small_tx(c, x, cy, r=3.6):
    c.circle(x, cy + r * 0.6, r)
    c.circle(x, cy - r * 0.6, r)
    return cy + r * 1.6, cy - r * 1.6


# ---------------------------------------------------------------- drawing
def build(res, p):
    c = Canvas()
    fmt_cb = lambda key: f"{res[key + '_cb']} A"
    isc33 = f"{res['Isc33_rated']:g} kA / {p['t_k']:g} s"
    isc11 = f"{res['Isc11_rated']:g} kA / {p['t_k']:g} s"

    # frame, zones, title block
    c.rect(5, 5, 410, 287, "FRAME", center=False)
    c.line(5, 42, 415, 42, layer="FRAME")
    c.line(250, 5, 250, 42, layer="FRAME")
    c.text(10, 270, "33 kV OUTDOOR", 3.2, bold=True)
    c.text(10, 266, "SWITCHYARD", 3.2, bold=True)
    c.rect(15, 47, 390, 113, "DASH", center=False)
    c.text(18, 153, "11 kV SWITCHGEAR ROOM", 2.2, bold=True)

    # 33 kV and 11 kV buses with couplers
    y33, y11 = 240, 134
    c.line(30, y33, 190, y33, "BUS")
    c.line(230, y33, 390, y33, "BUS")
    chain(c, "h", y33, 190, 230, [(197, "isolator"), (210, "breaker"), (223, "isolator")], "BUS")
    c.text(210, y33 + 6.5, "33 kV BUS COUPLER (N/O)", 2.1, ha="center", bold=True)
    c.text(210, y33 - 7.5, f"VCB {fmt_cb('cpl33')}, {isc33}", 1.9, ha="center")
    c.text(32, y33 + 3, "33 kV BUS", 2.6, bold=True)

    c.line(28, y11, 190, y11, "BUS")
    c.line(230, y11, 392, y11, "BUS")
    chain(c, "h", y11, 190, 230, [(210, "breaker")], "BUS")
    c.text(210, y11 + 6.5, "11 kV BUS COUPLER (N/O)", 2.1, ha="center", bold=True)
    c.text(210, y11 - 7.5, f"VCB {fmt_cb('cpl11')}, {isc11}", 1.9, ha="center")
    c.text(32, y11 + 3, "11 kV BUS", 2.6, bold=True)

    for half in (1, -1):  # +1 = left half, -1 = right half (mirror image)
        mx = (lambda x: x) if half == 1 else (lambda x: 420 - x)
        side_in = "left" if half == -1 else "right"  # labels toward the centre
        ha_in = "left" if half == 1 else "right"
        lab = lambda x, y, s, h=1.9: c.text(x + half * 5, y, s, h, ha=ha_in)
        n = 1 if half == 1 else 2

        # ---- 33 kV incomer
        X = mx(100)
        c.text(X, 289, f"33 kV INCOMER-{n}", 2.3, ha="center", bold=True)
        c.text(X, 285.5, "FROM 132/33 kV GRID SUBSTATION", 1.9, ha="center")
        arrow_down(c, X, 279)
        chain(c, "v", X, 279, y33, [(268, "isolator"), (261, "ct"), (254, "breaker"), (247, "isolator")])
        surge_arrester(c, X, 274, -half)
        lab(X, 262.5, f"CT {res['inc33_ct']}/1 A")
        lab(X, 255.5, f"VCB {fmt_cb('inc33')}")
        lab(X, 252.0, isc33)

        # ---- 33 kV bus PT
        XP = mx(160)
        top, bot = small_tx(c, XP, 226)
        c.line(XP, y33, XP, top)
        c.line(XP, bot, XP, bot - 0.1)
        earth(c, XP, bot)
        c.text(XP + half * 5, 225, "33 kV BUS PT", 1.9, ha=ha_in)

        # ---- transformer bay
        top_y, bot_y = power_tx(c, X, 185)
        chain(c, "v", X, y33, top_y, [(231, "isolator"), (222, "breaker"), (214, "ct")])
        surge_arrester(c, X, 209, half)
        lab(X, 223.5, f"VCB {fmt_cb('txhv')}")
        lab(X, 220.0, isc33)
        lab(X, 212.5, f"CT {res['txhv_ct']}/1 A")
        # neutral earthing resistor
        nx = X + half * 13
        c.line(X, 181, nx, 181)
        c.line(nx, 181, nx, 177)
        c.rect(nx, 173.5, 3, 7)
        c.line(nx, 170, nx, 167)
        earth(c, nx, 167)
        c.text(nx + half * 3, 172.5, "NGR", 1.9, ha="left" if half == 1 else "right")
        # transformer data block (outer side)
        tx_ha = "right" if half == 1 else "left"
        tx0 = X - half * 10
        c.text(tx0, 192, f"T{n}", 3.0, ha=tx_ha, bold=True)
        c.text(tx0, 188, f"{p['V_hv']:g}/{p['V_lv']:g} kV, {p['S_tx']:g} MVA", 2.0, ha=tx_ha)
        c.text(tx0, 184.5, "ONAN, Dyn11", 2.0, ha=tx_ha)
        c.text(tx0, 181, f"Z = {p['Z_tx'] * 100:.1f} %", 2.0, ha=tx_ha)

        # ---- 11 kV incomer
        chain(c, "v", X, bot_y, y11, [(155, "ct"), (146, "breaker")])
        lab(X, 156.5, f"CT {res['inc11_ct']}/1 A")
        lab(X, 147.5, f"VCB {fmt_cb('inc11')}")
        lab(X, 144.0, isc11)

        # ---- 11 kV bays below the bus
        if half == 1:
            feeder_x = [(38, 1), (68, 2), (98, 3)]
            pt_x, sat_x = 128, 158
        else:
            feeder_x = [(322, 4), (352, 5), (382, 6)]
            pt_x, sat_x = 292, 262
        for fx, i in feeder_x:
            chain(c, "v", fx, y11, 100, [(126, "ct"), (117, "breaker")])
            arrow_down(c, fx, 96.4)
            name, mva = cfg.FEEDERS[i - 1]
            c.text(fx - 3.2, 92, name.upper(), 1.9, rot=270)
            c.text(fx + 0.2, 92, f"{mva:.1f} MVA  VCB {res[f'fdr{i}_cb']} A  CT {res[f'fdr{i}_ct']}/1", 1.9, rot=270)
        # bus PT
        top, bot = small_tx(c, pt_x, 114, 3.4)
        c.line(pt_x, y11, pt_x, top)
        earth(c, pt_x, bot)
        c.text(pt_x - 3.2, bot - 8, "11 kV BUS PT", 1.9, rot=270)
        # station auxiliary transformer behind a fuse
        top, bot = small_tx(c, sat_x, 112, 3.6)
        chain(c, "v", sat_x, y11, top, [(126, "fuse")])
        c.line(sat_x, bot, sat_x, 100)
        arrow_down(c, sat_x, 96.4)
        c.text(sat_x - 3.2, 92, "STATION AUX TX", 1.9, rot=270)
        c.text(sat_x + 0.2, 92, f"11/0.415 kV  {p['sat_kva']:g} kVA", 1.9, rot=270)

    # ---- notes and legend (bottom left)
    c.text(10, 37, "NOTES", 2.4, bold=True)
    notes = [
        "1. Learning / portfolio exercise. NOT FOR CONSTRUCTION. Inputs are assumptions (see workbook sheet 'Inputs').",
        "2. Both bus couplers normally open: each transformer feeds its own 11 kV section. Coupler closes on loss of one source.",
        "3. 11 kV transformer neutrals earthed through NGR (assumed; confirm with utility practice).",
        "4. Ratings are calculated in the Excel workbook (Short_Circuit and Equipment_Ratings sheets).",
    ]
    for i, s in enumerate(notes):
        c.text(10, 33.2 - i * 3.6, s, 1.9)
    lx, ly = 18, 12
    legend = [("breaker", "VCB"), ("isolator", "ISOLATOR"), ("ct", "CT"), ("tx", "TRANSFORMER"),
              ("pt", "PT"), ("earth", "EARTH"), ("fuse", "FUSE")]
    for kind, label in legend:
        if kind in ("breaker", "isolator", "ct", "fuse"):
            c.line(lx - 6, ly, lx - 3.2, ly) if kind != "ct" else None
            sym(c, kind, "h", lx, ly)
            c.line(lx + 3.2, ly, lx + 6, ly) if kind != "ct" else None
        elif kind == "tx":
            c.circle(lx - 1.5, ly, 3.2)
            c.circle(lx + 1.5, ly, 3.2)
        elif kind == "pt":
            c.circle(lx - 1.2, ly, 2.4)
            c.circle(lx + 1.2, ly, 2.4)
        elif kind == "earth":
            earth(c, lx, ly + 3)
        c.text(lx + 8, ly - 0.8, label, 1.9)
        lx += 34

    # ---- title block
    c.text(255, 35, "33/11 kV SUBSTATION", 5.0, bold=True)
    c.text(255, 29, "SINGLE LINE DIAGRAM", 3.8, bold=True)
    c.text(255, 24, f"Design basis: {p['n_tx']} x {p['S_tx']:g} MVA, {p['V_hv']:g}/{p['V_lv']:g} kV, Dyn11, "
                    f"{len(cfg.FEEDERS)} x 11 kV feeders", 2.0)
    c.text(255, 20, f"33 kV fault level (assumed): {p['S_src']:g} MVA = {res['Ik33']:.1f} kA", 2.0)
    c.text(255, 16, "Drawing no.: SLD-01    Rev: A    Scale: NTS    Sheet 1 of 1", 2.0)
    c.text(255, 12, f"Prepared by: {cfg.AUTHOR}    Date: {datetime.date.today().isoformat()}", 2.0)
    c.text(255, 8, "LEARNING / PORTFOLIO EXERCISE - NOT FOR CONSTRUCTION", 2.0, bold=True)
    return c


# ---------------------------------------------------------------- renderers
def to_dxf(c, path):
    doc = ezdxf.new("R2010", setup=True)
    doc.units = ezdxf.units.MM
    for name, (col, lw, dashed) in LAYERS.items():
        doc.layers.add(name, color=col, lineweight=int(lw * 100), linetype="DASHED" if dashed else "Continuous")
    msp = doc.modelspace()
    align = {"left": TextEntityAlignment.LEFT, "center": TextEntityAlignment.CENTER, "right": TextEntityAlignment.RIGHT}
    for kind, layer, d in c.items:
        attr = {"layer": layer}
        if kind == "line":
            msp.add_line((d[0], d[1]), (d[2], d[3]), dxfattribs=attr)
        elif kind == "circle":
            msp.add_circle((d[0], d[1]), d[2], dxfattribs=attr)
        elif kind == "poly":
            pts, closed, fill = d
            if fill:
                msp.add_solid([pts[0], pts[1], pts[2], pts[2]], dxfattribs=attr)
            else:
                msp.add_lwpolyline(pts, close=closed, dxfattribs=attr)
        elif kind == "text":
            x, y, s, h, rot, ha, _bold = d
            t = msp.add_text(s, height=h, rotation=rot, dxfattribs=attr)
            t.set_placement((x, y), align=align[ha])
    doc.saveas(path)


def to_mpl(c, png, pdf):
    fig = plt.figure(figsize=(W / 25.4, H / 25.4))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W)
    ax.set_ylim(0, H)
    ax.axis("off")
    pt_per_mm = 72 / 25.4
    for kind, layer, d in c.items:
        _col, lw, dashed = LAYERS[layer]
        style = dict(color="black", linewidth=lw * pt_per_mm, solid_capstyle="butt")
        if dashed:
            style["linestyle"] = (0, (6, 3))
            style["color"] = "#666666"
        if kind == "line":
            ax.plot([d[0], d[2]], [d[1], d[3]], **style)
        elif kind == "circle":
            ax.add_patch(Circle((d[0], d[1]), d[2], fill=False, edgecolor="black", linewidth=lw * pt_per_mm))
        elif kind == "poly":
            pts, closed, fill = d
            ax.add_patch(Polygon(pts, closed=closed, fill=fill, facecolor="black" if fill else "none",
                                 edgecolor="#666666" if dashed else "black", linewidth=lw * pt_per_mm,
                                 linestyle="--" if dashed else "-"))
        elif kind == "text":
            x, y, s, h, rot, ha, bold = d
            ax.text(x, y, s, fontsize=h * 3.94, rotation=rot, ha=ha, va="baseline", rotation_mode="anchor",
                    fontweight="bold" if bold else "normal", family="DejaVu Sans", color="black")
    fig.savefig(png, dpi=150)
    fig.savefig(pdf)
    plt.close(fig)


def main(out_dir="drawings"):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    p = cfg.defaults()
    res = calcs.compute()
    c = build(res, p)
    base = out / "SLD_33_11kV_substation"
    to_dxf(c, str(base) + ".dxf")
    to_mpl(c, str(base) + ".png", str(base) + ".pdf")
    print("Wrote", base.with_suffix(".dxf"), base.with_suffix(".png"), base.with_suffix(".pdf"))


if __name__ == "__main__":
    main()
