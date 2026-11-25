"""Build the Excel design workbook (live formulas) from design_config.py.

Usage (from the project root):
    python src/build_workbook.py            # writes Substation_Design_Calculations.xlsx
Then open it in Excel, or let LibreOffice calculate it:
    python src/verify_workbook.py           # recalculates and cross-checks against calcs.py

Colour code used in every sheet:
    blue text         input typed by the designer
    blue on yellow    ASSUMPTION - replace with real data before relying on the result
    green text        value linked from another sheet
    black text        formula
"""
import datetime
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, ScatterChart, Series
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

import design_config as cfg

FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="008000")
F_TITLE = Font(name=FONT, size=16, bold=True, color="1F3864")
F_SUB = Font(name=FONT, size=10, italic=True, color="595959")
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
FILL_ASSUMED = PatternFill("solid", fgColor="FFF2CC")
FILL_RESULT = PatternFill("solid", fgColor="E2EFDA")
FILL_PASS = PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE")
FILL_FAIL = PatternFill("solid", bgColor="FFC7CE", fgColor="FFC7CE")
FILL_REVIEW = PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

REG = {}      # key -> absolute cell reference including sheet name
RNG = {}      # key -> absolute range reference


class _Fmt(dict):
    def __missing__(self, key):
        raise KeyError(f"Unknown key used in formula: {key}")


def sub(template):
    """Replace {key} placeholders with absolute cell references."""
    return template.format_map(_Fmt({**REG, **RNG}))


def NEXT(lst, x):
    """Excel text: smallest value in an ascending list that is >= x."""
    return f'IFERROR(INDEX({{{lst}}},COUNTIF({{{lst}}},"<"&({x}))+1),"> list")'


class SW:
    """Sheet writer for label / value / unit / note layouts."""

    def __init__(self, ws, widths, note_col=4):
        self.ws, self.r, self.note_col = ws, 1, note_col
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + i)].width = w
        ws.sheet_view.showGridLines = False

    def title(self, text, sub_text=None):
        c = self.ws.cell(self.r, 1, text)
        c.font = F_TITLE
        self.r += 1
        if sub_text:
            self.ws.cell(self.r, 1, sub_text).font = F_SUB
            self.r += 1
        self.r += 1

    def section(self, text, ncols=4):
        for col in range(1, ncols + 1):
            self.ws.cell(self.r, col).fill = FILL_SEC
        self.ws.cell(self.r, 1, text).font = F_BOLD
        self.r += 1

    def header(self, labels):
        for i, lab in enumerate(labels, start=1):
            c = self.ws.cell(self.r, i, lab)
            c.font, c.fill, c.border = F_HDR, FILL_HDR, BORDER
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        self.r += 1

    def gap(self, n=1):
        self.r += n

    def row(self, key, label, value, unit="", note="", fmt=None, kind="calc"):
        ws, r = self.ws, self.r
        ws.cell(r, 1, label).font = F_BASE
        if isinstance(value, str) and value.startswith("="):
            value = sub(value)
        c = ws.cell(r, 2, value)
        c.border = BORDER
        c.alignment = Alignment(horizontal="right")
        c.font = {"input": F_INPUT, "assumed": F_INPUT, "link": F_LINK, "result": F_BOLD}.get(kind, F_BASE)
        if kind == "assumed":
            c.fill = FILL_ASSUMED
        if kind == "result":
            c.fill = FILL_RESULT
        if fmt:
            c.number_format = fmt
        ws.cell(r, 3, unit).font = F_BASE
        n = ws.cell(r, self.note_col, note)
        n.font = F_SUB
        if key:
            REG[key] = f"{ws.title}!$B${r}"
        self.r += 1
        return r


def status_formatting(ws, cell_range):
    ws.conditional_formatting.add(cell_range, CellIsRule(operator="equal", formula=['"PASS"'], fill=FILL_PASS))
    ws.conditional_formatting.add(cell_range, CellIsRule(operator="equal", formula=['"OK"'], fill=FILL_PASS))
    ws.conditional_formatting.add(cell_range, CellIsRule(operator="equal", formula=['"FAIL"'], fill=FILL_FAIL))
    ws.conditional_formatting.add(cell_range, CellIsRule(operator="equal", formula=['"AUGMENT"'], fill=FILL_FAIL))
    ws.conditional_formatting.add(cell_range, CellIsRule(operator="equal", formula=['"REVIEW"'], fill=FILL_REVIEW))
    ws.conditional_formatting.add(cell_range, CellIsRule(operator="equal", formula=['"WATCH"'], fill=FILL_REVIEW))


def style_chart(ch, title, xt, yt):
    ch.title, ch.roundedCorners = title, False
    ch.x_axis.title, ch.y_axis.title = xt, yt
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.legend.overlay = False
    ch.title.overlay = False
    ch.x_axis.title.overlay = False
    ch.y_axis.title.overlay = False
    ch.height, ch.width = 9, 18


# ---------------------------------------------------------------------------------------------
def build(sld_png=None):
    REG.clear()
    RNG.clear()
    wb = Workbook()
    names = ["Cover", "Summary", "Inputs", "Lists", "Load_Growth", "Short_Circuit", "Equipment_Ratings",
             "Busbar", "Earthing", "Protection", "BOM", "SLD"]
    sheets = {}
    for i, n in enumerate(names):
        sheets[n] = wb.active if i == 0 else wb.create_sheet(n)
        sheets[n].title = n
    nf = len(cfg.FEEDERS)

    # ===================================================================== Lists
    ws = sheets["Lists"]
    ws.sheet_view.showGridLines = False
    ws.cell(1, 1, "Standard value lists and material data").font = F_TITLE
    ws.cell(2, 1, "Ascending lists used by the 'next standard size' formulas. Confirm against the utility "
                  "specification / standards before real use.").font = F_SUB

    def put_list(col, header, values, key, fmt=None, start=5):
        c = ws.cell(4, col, header)
        c.font, c.fill, c.border = F_HDR, FILL_HDR, BORDER
        c.alignment = Alignment(wrap_text=True, horizontal="center")
        for i, v in enumerate(values):
            cell = ws.cell(start + i, col, v)
            cell.font, cell.border = F_INPUT, BORDER
            if fmt:
                cell.number_format = fmt
        L = chr(64 + col)
        RNG[key] = f"Lists!${L}${start}:${L}${start + len(values) - 1}"
        ws.column_dimensions[L].width = 18

    put_list(1, "Short-circuit rating (kA)", cfg.STD_ISC_KA, "L_isc")
    put_list(2, "Rated current (A)", cfg.STD_RATED_A, "L_cur")
    put_list(3, "CT primary (A)", cfg.STD_CT_PRIMARY_A, "L_ct")
    put_list(4, "ACSR name", [n for n, _ in cfg.ACSR], "L_acsr_n")
    put_list(5, "ACSR nominal Al area (mm2)", [a for _, a in cfg.ACSR], "L_acsr_a")
    put_list(6, "Earth flat size (mm)", [n for n, _ in cfg.FLATS], "L_fl_n")
    put_list(7, "Earth flat area (mm2)", [a for _, a in cfg.FLATS], "L_fl_a")
    put_list(8, "Earthing material", [m[0] for m in cfg.EARTH_MATERIALS], "L_m_n")
    put_list(9, "alpha_r (1/deg C)", [m[1] for m in cfg.EARTH_MATERIALS], "L_m_ar", "0.00000")
    put_list(10, "K0 (deg C)", [m[2] for m in cfg.EARTH_MATERIALS], "L_m_k0")
    put_list(11, "Tm fusing temp (deg C)", [m[3] for m in cfg.EARTH_MATERIALS], "L_m_tm")
    put_list(12, "rho_r (uohm.cm)", [m[4] for m in cfg.EARTH_MATERIALS], "L_m_rho")
    put_list(13, "TCAP (J/cm3.degC)", [m[5] for m in cfg.EARTH_MATERIALS], "L_m_tcap")
    ws.column_dimensions["H"].width = 24
    ws.row_dimensions[4].height = 42
    ws.cell(14, 1, "Sources: preferred short-circuit and current ratings follow the IEC 62271 style series; ACSR nominal "
                   "areas are the usual designations; material constants follow IEEE Std 80 Table 2 style data. "
                   "Verify before use.").font = F_SUB

    # ===================================================================== Inputs
    ws = sheets["Inputs"]
    s = SW(ws, [50, 18, 14, 90])
    s.title("Design inputs", "Change only the blue cells. Yellow cells are assumptions: replace them with real data.")
    for sec, items in cfg.SECTIONS:
        s.section(sec)
        for key, label, value, unit, note, kind, fmt in items:
            s.row(key, label, value, unit, note, fmt, "assumed" if kind == "assumed" else "input")
    s.section("8. Outgoing 11 kV feeders (peak demand, ASSUMED illustrative loads)")
    for i, (name, mva) in enumerate(cfg.FEEDERS, start=1):
        s.row(f"f{i}", name, mva, "MVA", "ASSUMED peak load of this feeder", "0.00", "assumed")
    dv = DataValidation(type="list", formula1="=" + RNG["L_m_n"], allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(REG["mat"].split("!")[1].replace("$", ""))
    ws.freeze_panes = "A4"

    # ===================================================================== Load_Growth
    ws = sheets["Load_Growth"]
    s = SW(ws, [34, 14, 14, 14, 14, 16, 16, 12, 10, 10])
    s.title("Load growth and transformer capacity (N-1)",
            "Does the station still supply its load when one transformer is out of service?")
    s.section("Year-0 demand build-up", 4)
    for i, (name, _mva) in enumerate(cfg.FEEDERS, start=1):
        s.row(f"lg_f{i}", name, f"={{f{i}}}", "MVA", "", "0.00", "link")
    s.row("sum_feeders", "Sum of feeder peak loads", f"=SUM(B{s.r - nf}:B{s.r - 1})", "MVA", "", "0.00")
    s.row("D0", "Station peak demand, year 0", "={sum_feeders}*{diversity}", "MVA", "Sum of peaks x diversity factor", "0.00", "result")
    s.row("D0_MW", "Station peak demand, year 0", "={D0}*{pf}", "MW", "", "0.00")
    s.row("cap_inst", "Installed transformer capacity", "={n_tx}*{S_tx}", "MVA", "", "0.0")
    s.row("cap_n1", "Firm capacity with one transformer out (N-1)", "=({n_tx}-1)*{S_tx}*(1+{ov_n1})", "MVA", "", "0.0")
    s.row("load_norm0", "Normal loading, year 0", "={D0}/{cap_inst}", "% of installed", "", "0.0%")
    s.row("load_n10", "N-1 loading, year 0", "={D0}/{cap_n1}", "% of firm", "", "0.0%", "result")
    s.gap()
    s.section("Year-by-year check", 10)
    s.header(["Year", "Demand (MVA)", "Demand (MW)", "Normal loading", "N-1 loading", "Firm capacity (MVA)",
              "Normal limit (MVA)", "Status", "N-1 flag", "Limit flag"])
    y0 = s.r
    for n in range(0, cfg.N_YEARS + 1):
        r = s.r
        vals = [n, "={D0}*(1+{growth})^A%d" % r, "=B%d*{pf}" % r, "=B%d/{cap_inst}" % r, "=B%d/{cap_n1}" % r,
                "={cap_n1}", "={lim_norm}*{cap_inst}",
                '=IF(E%d>1,"AUGMENT",IF(D%d>{lim_norm},"WATCH","OK"))' % (r, r),
                "=IF(E%d>1,1,0)" % r, "=IF(D%d>{lim_norm},1,0)" % r]
        fmts = ["0", "0.00", "0.00", "0.0%", "0.0%", "0.0", "0.0", "@", "0", "0"]
        for i, (v, f) in enumerate(zip(vals, fmts), start=1):
            c = ws.cell(r, i, sub(v) if isinstance(v, str) else v)
            c.font, c.border, c.number_format = F_BASE, BORDER, f
        s.r += 1
    y1 = s.r - 1
    status_formatting(ws, f"H{y0}:H{y1}")
    s.gap()
    REG_sum_row = s.r
    s.row("breach_year", "First year N-1 capacity is exceeded",
          f'=IFERROR(INDEX(A{y0}:A{y1},MATCH(1,I{y0}:I{y1},0)),">15")', "year", "Transformer augmentation needed by then",
          "0", "result")
    s.row("watch_year", "First year normal loading exceeds the limit",
          f'=IFERROR(INDEX(A{y0}:A{y1},MATCH(1,J{y0}:J{y1},0)),">15")', "year", "", "0")
    ch = LineChart()
    style_chart(ch, "Station demand vs transformer capacity", "Year", "MVA")
    for col, name in ((2, "Demand"), (6, "Firm capacity (N-1)"), (7, "Normal loading limit")):
        ch.add_data(Reference(ws, min_col=col, min_row=y0 - 1, max_row=y1), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=y0, max_row=y1))
    ws.add_chart(ch, f"A{s.r + 2}")

    # ===================================================================== Short_Circuit
    ws = sheets["Short_Circuit"]
    s = SW(ws, [54, 16, 12, 90])
    s.title("Short-circuit study", "Balanced three-phase fault, per-unit method. Motor contribution and cable impedance ignored (conservative).")
    s.section("Impedances on the calculation base")
    s.row("Zs", "Source impedance", "={S_base}/{S_src}", "pu", "Base MVA / source fault MVA", "0.0000")
    s.row("Rs", "Source resistance", "={Zs}/SQRT(1+{XR_src}^2)", "pu", "", "0.0000")
    s.row("Xs", "Source reactance", "={Rs}*{XR_src}", "pu", "", "0.0000")
    s.row("Zt", "Transformer impedance (each)", "={Z_tx}*{S_base}/{S_tx}", "pu", "%Z x base MVA / transformer MVA", "0.0000")
    s.row("Rt", "Transformer resistance (each)", "={Zt}/SQRT(1+{XR_tx}^2)", "pu", "", "0.0000")
    s.row("Xt", "Transformer reactance (each)", "={Rt}*{XR_tx}", "pu", "", "0.0000")
    s.gap()
    s.section("A. Fault on 11 kV bus, ONE transformer in service (bus coupler open)")
    s.row("R_a", "Total resistance", "={Rs}+{Rt}", "pu", "", "0.0000")
    s.row("X_a", "Total reactance", "={Xs}+{Xt}", "pu", "", "0.0000")
    s.row("Z_a", "Total impedance", "=SQRT({R_a}^2+{X_a}^2)", "pu", "", "0.0000")
    s.row("Ik_a", "Initial symmetrical fault current Ik''", "={S_base}/(SQRT(3)*{V_lv}*{Z_a})", "kA", "", "0.000", "result")
    s.row("S_a", "Fault level", "={S_base}/{Z_a}", "MVA", "", "0.0")
    s.row("XR_a", "X/R at fault", "={X_a}/{R_a}", "-", "", "0.0")
    s.row("k_a", "Peak factor kappa", "=1.02+0.98*EXP(-3*{R_a}/{X_a})", "-", "IEC 60909 style", "0.000")
    s.row("ip_a", "Peak make current ip", "={k_a}*SQRT(2)*{Ik_a}", "kA", "", "0.00")
    s.gap()
    s.section("B. Fault on 11 kV bus, ALL transformers in parallel (bus coupler closed)")
    s.row("R_b", "Total resistance", "={Rs}+{Rt}/{n_tx}", "pu", "", "0.0000")
    s.row("X_b", "Total reactance", "={Xs}+{Xt}/{n_tx}", "pu", "", "0.0000")
    s.row("Z_b", "Total impedance", "=SQRT({R_b}^2+{X_b}^2)", "pu", "", "0.0000")
    s.row("Ik_b", "Initial symmetrical fault current Ik''", "={S_base}/(SQRT(3)*{V_lv}*{Z_b})", "kA", "Worst case for 11 kV switchgear", "0.000", "result")
    s.row("S_b", "Fault level", "={S_base}/{Z_b}", "MVA", "", "0.0")
    s.row("XR_b", "X/R at fault", "={X_b}/{R_b}", "-", "", "0.0")
    s.row("k_b", "Peak factor kappa", "=1.02+0.98*EXP(-3*{R_b}/{X_b})", "-", "", "0.000")
    s.row("ip_b", "Peak make current ip", "={k_b}*SQRT(2)*{Ik_b}", "kA", "", "0.00")
    s.gap()
    s.section("C. Fault on 33 kV bus")
    s.row("Ik33", "Initial symmetrical fault current Ik''", "={S_src}/(SQRT(3)*{V_hv})", "kA", "From the assumed source fault level", "0.000", "result")
    s.row("k33", "Peak factor kappa", "=1.02+0.98*EXP(-3/{XR_src})", "-", "", "0.000")
    s.row("ip33", "Peak make current ip", "={k33}*SQRT(2)*{Ik33}", "kA", "", "0.00")
    s.gap()
    s.section("Switchgear short-circuit ratings selected")
    s.row("Ik11", "Design fault current, 11 kV", "=MAX({Ik_a},{Ik_b})", "kA", "Larger of cases A and B", "0.000")
    s.row("Isc33_rated", "33 kV switchgear rated short-circuit current",
          "=" + NEXT("L_isc", "{Ik33}*(1+{m_isc})"), "kA", "Next standard value above Ik'' plus margin", "0.0", "result")
    s.row("Isc11_rated", "11 kV switchgear rated short-circuit current",
          "=" + NEXT("L_isc", "{Ik11}*(1+{m_isc})"), "kA", "Utilities often standardise a higher value; check the spec", "0.0", "result")
    s.row("ip33_rated", "33 kV rated peak withstand", "=2.5*{Isc33_rated}", "kA", "2.5 x Isc for 50 Hz (IEC 62271)", "0.0")
    s.row("ip11_rated", "11 kV rated peak withstand", "=2.5*{Isc11_rated}", "kA", "", "0.0")
    s.row(None, "Rated short-time duration", "={t_k}", "s", "", "0", "link")

    # ===================================================================== Equipment_Ratings
    ws = sheets["Equipment_Ratings"]
    s = SW(ws, [46, 18, 18, 20, 16, 16, 14, 16])
    s.title("Equipment ratings", "Rated currents, CT ratios and fault ratings for each circuit. Rule: rated current = next standard value above (load current x (1 + margin)).")
    s.section("Transformer full-load currents", 8)
    s.row("I_hv_flc", "HV full-load current (33 kV side)", "={S_tx}*1000/(SQRT(3)*{V_hv})", "A", "", "0.0", "result")
    s.row("I_lv_flc", "LV full-load current (11 kV side)", "={S_tx}*1000/(SQRT(3)*{V_lv})", "A", "", "0.0", "result")
    s.gap()
    s.section("Circuit ratings", 8)
    s.header(["Circuit", "Load current basis (A)", "With margin (A)", "Breaker / isolator rated current (A)",
              "CT primary (A)", "Fault rating (kA)", "Duration (s)", "Peak withstand (kA)"])
    circuits = [
        ("inc33", "33 kV incomer (one incomer carries both transformers)", "={n_tx}*{I_hv_flc}", "33"),
        ("cpl33", "33 kV bus coupler", "={I_hv_flc}", "33"),
        ("txhv", "33 kV transformer HV bay", "={I_hv_flc}", "33"),
        ("bus33", "33 kV busbar", "={n_tx}*{I_hv_flc}", "33"),
        ("inc11", "11 kV transformer incomer", "={I_lv_flc}", "11"),
        ("cpl11", "11 kV bus coupler", "={I_lv_flc}", "11"),
        ("bus11", "11 kV busbar (in switchgear)", "={I_lv_flc}", "11"),
    ] + [(f"fdr{i}", f"11 kV feeder: {name}", f"={{f{i}}}*1000/(SQRT(3)*{{V_lv}})", "11")
         for i, (name, _m) in enumerate(cfg.FEEDERS, start=1)]
    first_row = s.r
    for key, label, basis, lvl in circuits:
        r = s.r
        isc = "{Isc33_rated}" if lvl == "33" else "{Isc11_rated}"
        cells = [
            (label, F_BASE, None),
            (sub(basis), F_BASE, "0.0"),
            (f"=B{r}*(1+{REG['m_i']})", F_BASE, "0.0"),
            ("=" + sub(NEXT("L_cur", f"C{r}")), F_BOLD, "0"),
            ("=" + sub(NEXT("L_ct", f"C{r}")), F_BOLD, "0"),
            ("=" + sub(isc), F_BASE, "0.0"),
            ("=" + sub("{t_k}"), F_BASE, "0"),
            (f"=2.5*F{r}", F_BASE, "0.0"),
        ]
        for i, (v, fnt, f) in enumerate(cells, start=1):
            c = ws.cell(r, i, v)
            c.font, c.border = fnt, BORDER
            if f:
                c.number_format = f
        for col, suffix in ((2, "basis"), (3, "design"), (4, "cb"), (5, "ct")):
            REG[f"{key}_{suffix}"] = f"Equipment_Ratings!${chr(64 + col)}${r}"
        s.r += 1
    last_row = s.r - 1
    fr = first_row + len(circuits) - nf
    RNG["fdr_basis"] = f"Equipment_Ratings!$B${fr}:$B${last_row}"
    s.gap()
    ws.cell(s.r, 1, "Notes: CT secondary assumed 1 A. 11 kV switchgear is type-tested metal-clad; the bus rating comes from the manufacturer "
                    "range. Many utilities standardise 1250 A bus / higher fault ratings for expansion: follow the utility specification.").font = F_SUB

    # ===================================================================== Busbar
    ws = sheets["Busbar"]
    s = SW(ws, [54, 16, 12, 90])
    s.title("33 kV busbar conductor", "Minimum area from the short-circuit thermal limit (adiabatic). Ampacity is NOT checked here.")
    s.section("Thermal withstand")
    s.row("k_al", "Material constant k for aluminium", "=SQRT({Qc}*({beta}+20)/{rho20}*LN(({beta}+{th_f})/({beta}+{th_i})))*1E-6",
          "A.s^0.5/mm2", "IEC 60949 adiabatic formula", "0.0")
    s.row("A_min_bus", "Minimum conductor area", "={Ik33}*1000*SQRT({t_bus})/{k_al}", "mm2", "A = Ik'' x sqrt(t) / k", "0.0", "result")
    s.row("A_sel_bus", "Selected ACSR nominal aluminium area", "=" + NEXT("L_acsr_a", "{A_min_bus}"), "mm2", "Next standard size", "0", "result")
    s.row("bus_name", "Selected ACSR conductor", '=IFERROR(INDEX({L_acsr_n},MATCH({A_sel_bus},{L_acsr_a},0)),"none")', "-", "", "@", "result")
    s.row("J_bus", "Current density at design current", "={bus33_design}/{A_sel_bus}", "A/mm2", "Informational. Check ampacity in the manufacturer data at site conditions.", "0.00")
    s.row("bus_check", "Thermal check", '=IF(ISNUMBER({A_sel_bus}),"PASS","FAIL")', "-", "FAIL means no listed size is large enough", "@", "result")
    status_formatting(ws, "B1:B30")
    s.gap()
    ws.cell(s.r, 1, "Not covered: ampacity, sag/tension, electrodynamic forces, corona, clearances. 11 kV busbars are inside factory-built switchgear.").font = F_SUB

    # ===================================================================== Earthing
    ws = sheets["Earthing"]
    s = SW(ws, [54, 16, 12, 90])
    s.title("Earthing grid design", "Simplified IEEE Std 80 method for a rectangular grid. Verify the soil model on site before any real design.")
    s.section("1. Conductor size")
    s.row("m_ar", "alpha_r of selected material", "=INDEX({L_m_ar},MATCH({mat},{L_m_n},0))", "1/deg C", "", "0.00000", "link")
    s.row("m_k0", "K0", "=INDEX({L_m_k0},MATCH({mat},{L_m_n},0))", "deg C", "", "0", "link")
    s.row("m_tm", "Tm (fusing temperature)", "=INDEX({L_m_tm},MATCH({mat},{L_m_n},0))", "deg C", "", "0", "link")
    s.row("m_rho", "rho_r", "=INDEX({L_m_rho},MATCH({mat},{L_m_n},0))", "uohm.cm", "", "0.00", "link")
    s.row("m_tcap", "TCAP", "=INDEX({L_m_tcap},MATCH({mat},{L_m_n},0))", "J/cm3.degC", "", "0.00", "link")
    s.row("A_req_earth", "Minimum conductor area (thermal)",
          "={I_3I0}/SQRT(({m_tcap}*1E-4/({t_c}*{m_ar}*{m_rho}))*LN(1+({m_tm}-{T_amb})/({m_k0}+{T_amb})))",
          "mm2", "IEEE 80 conductor sizing", "0.0")
    s.row("A_sel_earth", "Selected flat area", "=" + NEXT("L_fl_a", "MAX({A_req_earth},{min_area})"), "mm2",
          "Larger of thermal need and the assumed minimum size", "0", "result")
    s.row("flat_name", "Selected flat size", "=INDEX({L_fl_n},MATCH({A_sel_earth},{L_fl_a},0))", "mm", "", "@", "result")
    s.row("d_cond", "Equivalent round diameter", "=2*SQRT({A_sel_earth}/PI())/1000", "m", "Equal-area round conductor (conservative)", "0.0000")
    s.gap()
    s.section("2. Grid geometry and resistance")
    s.row("nA", "Conductors along the length", "=ROUND({Ly}/{D},0)+1", "nos", "", "0")
    s.row("nB", "Conductors along the width", "=ROUND({Lx}/{D},0)+1", "nos", "", "0")
    s.row("LC", "Total horizontal conductor length", "={nA}*{Lx}+{nB}*{Ly}", "m", "", "#,##0")
    s.row("Agrid", "Grid area", "={Lx}*{Ly}", "m2", "", "#,##0")
    s.row("LP", "Grid perimeter", "=2*({Lx}+{Ly})", "m", "", "0")
    s.row("LR", "Total rod length", "={rods_n}*{rods_L}", "m", "", "0")
    s.row("LT", "Total buried conductor length", "={LC}+{LR}", "m", "", "#,##0")
    s.row("Rg", "Grid resistance Rg", "={rho_soil}*(1/{LT}+1/SQRT(20*{Agrid})*(1+1/(1+{h}*SQRT(20/{Agrid}))))", "ohm",
          "Sverak formula", "0.000", "result")
    s.gap()
    s.section("3. Fault current into the grid")
    s.row("Ta_dc", "DC offset time constant", "={XR_g}/(2*PI()*{freq})", "s", "", "0.0000")
    s.row("Df", "Decrement factor", "=SQRT(1+({Ta_dc}/{t_f})*(1-EXP(-2*{t_f}/{Ta_dc})))", "-", "", "0.000")
    s.row("Ig", "Maximum grid current Ig", "={Df}*{Sf}*{I_3I0}*1000", "A", "", "#,##0")
    s.row("GPR", "Ground potential rise", "={Ig}*{Rg}", "V", "", "#,##0", "result")
    s.gap()
    s.section("4. Tolerable voltages")
    s.row("Cs", "Surface layer derating factor Cs", "=1-0.09*(1-{rho_soil}/{rho_s})/(2*{hs}+0.09)", "-", "", "0.000")
    s.row("Kb", "Body constant", "=IF({body}=70,0.157,0.116)", "-", "0.157 for 70 kg, 0.116 for 50 kg", "0.000")
    s.row("Etouch", "Tolerable touch voltage", "=(1000+1.5*{Cs}*{rho_s})*{Kb}/SQRT({t_f})", "V", "", "#,##0", "result")
    s.row("Estep", "Tolerable step voltage", "=(1000+6*{Cs}*{rho_s})*{Kb}/SQRT({t_f})", "V", "", "#,##0", "result")
    s.gap()
    s.section("5. Mesh and step voltage of the design")
    s.row("na", "Factor na", "=2*{LC}/{LP}", "-", "", "0.000")
    s.row("nb", "Factor nb", "=SQRT({LP}/(4*SQRT({Agrid})))", "-", "nc = nd = 1 for a rectangular grid", "0.000")
    s.row("n_eff", "Effective number of parallel conductors n", "={na}*{nb}", "-", "", "0.000")
    s.row("Ki", "Irregularity factor Ki", "=0.644+0.148*{n_eff}", "-", "", "0.000")
    s.row("Kh", "Depth factor Kh", "=SQRT(1+{h}/1)", "-", "", "0.000")
    s.row("Kii", "Corner factor Kii", "=IF({rods_perim}=1,1,1/(2*{n_eff})^(2/{n_eff}))", "-", "", "0.000")
    s.row("Km", "Mesh factor Km",
          "=(1/(2*PI()))*(LN({D}^2/(16*{h}*{d_cond})+({D}+2*{h})^2/(8*{D}*{d_cond})-{h}/(4*{d_cond}))+({Kii}/{Kh})*LN(8/(PI()*(2*{n_eff}-1))))",
          "-", "", "0.000")
    s.row("LM", "Effective length for mesh voltage", "={LC}+(1.55+1.22*({rods_L}/SQRT({Lx}^2+{Ly}^2)))*{LR}", "m", "", "#,##0.0")
    s.row("Em", "Mesh (touch) voltage of the design", "={rho_soil}*{Km}*{Ki}*{Ig}/{LM}", "V", "", "#,##0", "result")
    s.row("Ks", "Step factor Ks", "=(1/PI())*(1/(2*{h})+1/({D}+{h})+(1/{D})*(1-0.5^({n_eff}-2)))", "-", "", "0.000")
    s.row("LS", "Effective length for step voltage", "=0.75*{LC}+0.85*{LR}", "m", "", "#,##0.0")
    s.row("Es", "Step voltage of the design", "={rho_soil}*{Ks}*{Ki}*{Ig}/{LS}", "V", "", "#,##0", "result")
    s.gap()
    s.section("6. Checks")
    s.row("chk_cond", "Conductor area adequate", '=IF({A_sel_earth}>={A_req_earth},"PASS","FAIL")', "-", "", "@", "result")
    s.row("chk_touch", "Mesh voltage below tolerable touch voltage", '=IF({Em}<{Etouch},"PASS","FAIL")', "-", "", "@", "result")
    s.row("util_touch", "Mesh voltage as share of tolerable", "={Em}/{Etouch}", "%", "Lower is safer", "0%")
    s.row("chk_step", "Step voltage below tolerable step voltage", '=IF({Es}<{Estep},"PASS","FAIL")', "-", "", "@", "result")
    s.row("chk_gpr", "GPR below tolerable touch voltage", '=IF({GPR}<={Etouch},"PASS","REVIEW")', "-",
          "REVIEW: check transferred potentials (fence, LV neutral, telecom, pipes) per IEEE 80", "@", "result")
    status_formatting(ws, "B1:B80")

    # ===================================================================== Protection
    ws = sheets["Protection"]
    s = SW(ws, [54, 16, 12, 90])
    s.title("Overcurrent relay grading", "IEC standard-inverse IDMT curves graded at the maximum 11 kV fault current (one transformer in service).")
    s.section("Fault current for grading")
    s.row("If_prot", "11 kV fault current, one transformer", "={Ik_a}*1000", "A", "Case A of the short-circuit study", "#,##0")
    s.row("I_fdr_max", "Largest feeder load current", f"=MAX({RNG['fdr_basis']})", "A", "", "0.0")
    s.gap()
    s.section("11 kV feeder relay (downstream)")
    s.row("Ip_f", "Pickup current", "={pk_f}*{I_fdr_max}", "A", "", "0.0")
    s.row("PSM_f", "Plug setting multiple at fault", "={If_prot}/{Ip_f}", "-", "", "0.00")
    s.row("t_f_relay", "Operating time", "={tms_f}*{k_si}/({PSM_f}^{a_si}-1)", "s", "", "0.000", "result")
    s.gap()
    s.section("11 kV incomer relay")
    s.row("Ip_inc", "Pickup current", "={pk_inc}*{I_lv_flc}", "A", "", "0.0")
    s.row("PSM_inc", "Plug setting multiple at fault", "={If_prot}/{Ip_inc}", "-", "", "0.00")
    s.row("TMS_inc_req", "TMS needed for the grading margin", "=({t_f_relay}+{margin})*({PSM_inc}^{a_si}-1)/{k_si}", "-", "", "0.0000")
    s.row("TMS_inc", "TMS set (rounded up to relay step)", "=CEILING({TMS_inc_req},{tms_step})", "-", "", "0.000", "result")
    s.row("t_inc", "Operating time", "={TMS_inc}*{k_si}/({PSM_inc}^{a_si}-1)", "s", "", "0.000", "result")
    s.gap()
    s.section("33 kV transformer relay (upstream)")
    s.row("I_hv_fault", "Fault current seen on the 33 kV side", "={If_prot}*{V_lv}/{V_hv}", "A", "Three-phase fault, Dyn11 ratio", "#,##0")
    s.row("Ip_hv", "Pickup current", "={pk_hv}*{I_hv_flc}", "A", "", "0.0")
    s.row("PSM_hv", "Plug setting multiple at fault", "={I_hv_fault}/{Ip_hv}", "-", "", "0.00")
    s.row("TMS_hv_req", "TMS needed for the grading margin", "=({t_inc}+{margin})*({PSM_hv}^{a_si}-1)/{k_si}", "-", "", "0.0000")
    s.row("TMS_hv", "TMS set (rounded up to relay step)", "=CEILING({TMS_hv_req},{tms_step})", "-", "", "0.000", "result")
    s.row("t_hv", "Operating time", "={TMS_hv}*{k_si}/({PSM_hv}^{a_si}-1)", "s", "", "0.000", "result")
    s.gap()
    s.section("Checks")
    s.row("margin_inc", "Margin: incomer minus feeder", "={t_inc}-{t_f_relay}", "s", "", "0.000")
    s.row("margin_hv", "Margin: 33 kV relay minus incomer", "={t_hv}-{t_inc}", "s", "", "0.000")
    s.row("chk_grade", "Grading margins achieved", '=IF(AND({margin_inc}>={margin}-0.0001,{margin_hv}>={margin}-0.0001),"PASS","FAIL")', "-",
          "Only the maximum-fault case is checked. Minimum-fault sensitivity and instantaneous settings are not covered.", "@", "result")
    status_formatting(ws, "B1:B50")
    s.gap()
    s.section("Time-current curves (current referred to the 11 kV side)", 4)
    for col, lab in enumerate(["Fault current (A)", "Feeder relay (s)", "Incomer relay (s)", "33 kV relay (s)"], start=1):
        c = ws.cell(s.r, col, lab)
        c.font, c.fill, c.border = F_HDR, FILL_HDR, BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    s.r += 1
    c0 = s.r
    for amps in (1000, 1500, 2000, 3000, 4000, 5000, 6000, 8000, 10000):
        r = s.r
        vals = [amps,
                sub("={tms_f}*{k_si}/((A%d/{Ip_f})^{a_si}-1)" % r),
                sub("={TMS_inc}*{k_si}/((A%d/{Ip_inc})^{a_si}-1)" % r),
                sub("={TMS_hv}*{k_si}/(((A%d*{V_lv}/{V_hv})/{Ip_hv})^{a_si}-1)" % r)]
        for i, v in enumerate(vals, start=1):
            c = ws.cell(r, i, v)
            c.font, c.border = (F_INPUT if i == 1 else F_BASE), BORDER
            c.number_format = "#,##0" if i == 1 else "0.00"
        s.r += 1
    c1 = s.r - 1
    sc = ScatterChart()
    style_chart(sc, "Overcurrent relay curves", "Fault current (A, 11 kV side)", "Operating time (s)")
    sc.scatterStyle = "lineMarker"
    sc.x_axis.scaling.logBase = 10
    sc.y_axis.scaling.logBase = 10
    sc.y_axis.scaling.min, sc.y_axis.scaling.max = 0.1, 10
    xs = Reference(ws, min_col=1, min_row=c0, max_row=c1)
    for col in (2, 3, 4):
        se = Series(Reference(ws, min_col=col, min_row=c0 - 1, max_row=c1), xs, title_from_data=True)
        sc.series.append(se)
    ws.add_chart(sc, f"A{s.r + 2}")

    # ===================================================================== BOM
    ws = sheets["BOM"]
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFG", (6, 44, 62, 8, 8, 16, 18)):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["H"].width = 50
    ws.cell(1, 1, "Bill of quantities (equipment list)").font = F_TITLE
    ws.cell(2, 1, "Quantities follow the single line diagram. Unit rates are left blank on purpose: enter rates from the utility schedule of rates "
                  "or supplier quotations.").font = F_SUB
    hdr = ["No.", "Item", "Specification (from the design)", "Unit", "Qty", "Rate (INR)", "Amount (INR)", "Remarks"]
    for i, h in enumerate(hdr, start=1):
        c = ws.cell(4, i, h)
        c.font, c.fill, c.border = F_HDR, FILL_HDR, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    kvA = lambda key: '"33 kV, "&TEXT({%s_cb},"0")&" A, "&TEXT({Isc33_rated},"0.0")&" kA / "&TEXT({t_k},"0")&" s"' % key
    kvB = lambda key: '"11 kV, "&TEXT({%s_cb},"0")&" A, "&TEXT({Isc11_rated},"0.0")&" kA / "&TEXT({t_k},"0")&" s"' % key
    items = [
        ("SEC", "A. 33 kV outdoor switchyard"),
        ("Power transformer", '="33/11 kV, "&TEXT({S_tx},"0.0")&" MVA, ONAN, Dyn11, Z = "&TEXT({Z_tx},"0.0%")', "nos", "={n_tx}", "Dyn11, 11 kV neutral via NGR"),
        ("33 kV outdoor VCB - incomer", "=" + kvA("inc33"), "nos", "=2", ""),
        ("33 kV outdoor VCB - transformer bay", "=" + kvA("txhv"), "nos", "={n_tx}", ""),
        ("33 kV outdoor VCB - bus coupler", "=" + kvA("cpl33"), "nos", "=1", ""),
        ("33 kV isolators - incomer bays", '="33 kV, "&TEXT({inc33_cb},"0")&" A"', "nos", "=2*2", "Two isolators per incomer bay"),
        ("33 kV isolators - transformer bays", '="33 kV, "&TEXT({txhv_cb},"0")&" A"', "nos", "={n_tx}", ""),
        ("33 kV isolators - bus coupler", '="33 kV, "&TEXT({cpl33_cb},"0")&" A"', "nos", "=2", ""),
        ("33 kV CT sets - incomer", '="CT "&TEXT({inc33_ct},"0")&"/1 A"', "sets", "=2", "3 CTs per set"),
        ("33 kV CT sets - transformer bay", '="CT "&TEXT({txhv_ct},"0")&"/1 A"', "sets", "={n_tx}", ""),
        ("33 kV surge arrester sets", "33 kV class, rating per utility specification", "sets", "=2+{n_tx}", "Incomers and transformer HV"),
        ("33 kV bus PT sets", "33 kV, ratio and burden per specification", "sets", "=2", "One per bus section"),
        ("Busbar and jumper conductor", '="ACSR "&{bus_name}&" ("&TEXT({A_sel_bus},"0")&" mm2 Al)"', "lot", "=1", "Length from the yard layout"),
        ("Neutral earthing resistor (NGR)", "11 kV neutral, ohmic value and current per earth-fault study", "nos", "={n_tx}", "Value not calculated here"),
        ("SEC", "B. 11 kV switchgear room"),
        ("11 kV VCB panel - transformer incomer", "=" + kvB("inc11") + '&", CT "&TEXT({inc11_ct},"0")&"/1 A"', "nos", "={n_tx}", ""),
        ("11 kV VCB panel - bus coupler", "=" + kvB("cpl11"), "nos", "=1", "Plus one bus riser panel (not counted)"),
        ("11 kV VCB panel - outgoing feeder", '="11 kV, "&TEXT(MAX(%s),"0")&" A VCB, CT ratios per feeder (see Equipment_Ratings)"' % "{fdr1_cb}:{fdr6_cb}".replace("{fdr1_cb}", "Equipment_Ratings!D%d" % fr).replace("{fdr6_cb}", "D%d" % last_row), "nos", "=%d" % nf, "Feeder breaker rating from Equipment_Ratings"),
        ("11 kV bus PT panel", "11 kV, ratio and burden per specification", "nos", "=2", "One per bus section"),
        ("Station auxiliary transformer with fuse switch", '="11/0.415 kV, "&TEXT({sat_kva},"0")&" kVA"', "nos", "=2", ""),
        ("SEC", "C. Auxiliary, earthing and civil"),
        ("DC supply (battery and charger)", "Rating per protection and control load", "lot", "=1", "Not sized in this package"),
        ("Control and relay panels", "Overcurrent and earth-fault relays per bay", "lot", "=1", "Settings in Protection sheet"),
        ("Earthing grid conductor", '={flat_name}&" mm "&{mat}&" flat"', "m", "={LC}", "From the Earthing sheet"),
        ("Earth rods", '=TEXT({rods_L},"0.0")&" m long rods"', "nos", "={rods_n}", ""),
        ("Surface layer (crushed rock)", '=TEXT({hs}*1000,"0")&" mm thick layer over the grid area"', "m3", "={Agrid}*{hs}", "Grid area used as an approximation of the yard"),
        ("Structures, cables, civil works", "Not estimated in this package", "lot", "=1", "Needs a site layout and the utility design standards"),
    ]
    r = 5
    no = 0
    first_item = r
    for it in items:
        if it[0] == "SEC":
            for col in range(1, 9):
                ws.cell(r, col).fill = FILL_SEC
            ws.cell(r, 1, it[1]).font = F_BOLD
            r += 1
            continue
        no += 1
        name, spec, unit, qty, remark = it
        ws.cell(r, 1, no).font = F_BASE
        ws.cell(r, 2, name).font = F_BASE
        ws.cell(r, 3, sub(spec) if spec.startswith("=") else spec).font = F_BASE
        ws.cell(r, 4, unit).font = F_BASE
        q = ws.cell(r, 5, sub(qty))
        q.font, q.number_format = F_BASE, "#,##0"
        rate = ws.cell(r, 6)
        rate.fill, rate.font, rate.number_format = FILL_ASSUMED, F_INPUT, "#,##0"
        amt = ws.cell(r, 7, f'=IF(F{r}="","",E{r}*F{r})')
        amt.font, amt.number_format = F_BASE, "#,##0"
        ws.cell(r, 8, remark).font = F_SUB
        for col in range(1, 8):
            ws.cell(r, col).border = BORDER
        r += 1
    last_item = r - 1
    ws.cell(r + 1, 6, "Total").font = F_BOLD
    t = ws.cell(r + 1, 7, f"=SUM(G{first_item}:G{last_item})")
    t.font, t.number_format = F_BOLD, "#,##0"
    ws.cell(r + 3, 1, "Rates are intentionally blank (yellow). No prices are assumed in this package.").font = F_SUB

    # ===================================================================== SLD
    ws = sheets["SLD"]
    ws.sheet_view.showGridLines = False
    ws.cell(1, 1, "Single line diagram").font = F_TITLE
    ws.cell(2, 1, "Same drawing as drawings/SLD_33_11kV_substation.dxf / .pdf. Ratings printed on it come from the Equipment_Ratings sheet.").font = F_SUB
    if sld_png and Path(sld_png).exists():
        img = XLImage(str(sld_png))
        scale = 1250 / img.width
        img.width, img.height = int(img.width * scale), int(img.height * scale)
        ws.add_image(img, "A4")

    # ===================================================================== Summary
    ws = sheets["Summary"]
    s = SW(ws, [56, 18, 14, 14, 70], note_col=5)
    s.title("Design summary", "Key results and checks. Everything here is a live link to the calculation sheets.")
    s.section("Key results", 5)
    rows = [
        ("Station peak demand, year 0", "={D0}", "MVA", "0.00"),
        ("N-1 loading, year 0", "={load_n10}", "%", "0.0%"),
        ("33 kV fault current", "={Ik33}", "kA", "0.00"),
        ("11 kV fault current (design, coupler closed)", "={Ik11}", "kA", "0.00"),
        ("33 kV switchgear fault rating selected", "={Isc33_rated}", "kA", "0.0"),
        ("11 kV switchgear fault rating selected", "={Isc11_rated}", "kA", "0.0"),
        ("33 kV incomer breaker rating", "={inc33_cb}", "A", "0"),
        ("11 kV transformer incomer breaker rating", "={inc11_cb}", "A", "0"),
        ("33 kV bus conductor", "={bus_name}", "-", "@"),
        ("Earth grid conductor", "={flat_name}", "mm", "@"),
        ("Grid resistance", "={Rg}", "ohm", "0.00"),
        ("Ground potential rise", "={GPR}", "V", "#,##0"),
        ("Mesh (touch) voltage of the design", "={Em}", "V", "#,##0"),
        ("Relay settings: feeder / incomer / 33 kV (TMS)", '=TEXT({tms_f},"0.000")&" / "&TEXT({TMS_inc},"0.000")&" / "&TEXT({TMS_hv},"0.000")', "-", "@"),
    ]
    for label, f, unit, fmt in rows:
        s.row(None, label, f, unit, "", fmt, "link")
    s.gap()
    s.section("Design checks", 5)
    s.header(["Check", "Result", "", "", "Detail"])
    checks = [
        ("N-1 loading in year 0 within firm capacity", '=IF({load_n10}<=1,"PASS","FAIL")', '="Loading "&TEXT({load_n10},"0%")'),
        ("Normal loading in year 0 within planning limit", '=IF({load_norm0}<={lim_norm},"PASS","FAIL")', '="Loading "&TEXT({load_norm0},"0%")&" vs limit "&TEXT({lim_norm},"0%")'),
        ("N-1 capacity lasts at least 5 years", '=IF(ISNUMBER({breach_year}),IF({breach_year}>=5,"PASS","REVIEW"),"PASS")', '="First breach: year "&{breach_year}&". Plan augmentation before then."'),
        ("Switchgear fault ratings found in standard list", '=IF(AND(ISNUMBER({Isc33_rated}),ISNUMBER({Isc11_rated})),"PASS","FAIL")', '="33 kV: "&TEXT({Isc33_rated},"0.0")&" kA, 11 kV: "&TEXT({Isc11_rated},"0.0")&" kA"'),
        ("33 kV bus conductor thermally adequate", "={bus_check}", '="Minimum "&TEXT({A_min_bus},"0")&" mm2, selected "&{bus_name}'),
        ("Earth conductor area adequate", "={chk_cond}", '="Required "&TEXT({A_req_earth},"0")&" mm2, selected "&TEXT({A_sel_earth},"0")&" mm2"'),
        ("Mesh voltage below tolerable touch voltage", "={chk_touch}", '=TEXT({Em},"#,##0")&" V vs "&TEXT({Etouch},"#,##0")&" V"'),
        ("Step voltage below tolerable step voltage", "={chk_step}", '=TEXT({Es},"#,##0")&" V vs "&TEXT({Estep},"#,##0")&" V"'),
        ("GPR below tolerable touch voltage", "={chk_gpr}", '="GPR "&TEXT({GPR},"#,##0")&" V: transferred potentials need review if above "&TEXT({Etouch},"#,##0")&" V"'),
        ("Relay grading margins achieved", "={chk_grade}", '="Margins "&TEXT({margin_inc},"0.00")&" s and "&TEXT({margin_hv},"0.00")&" s"'),
    ]
    c0 = s.r
    for label, f, detail in checks:
        r = s.r
        ws.cell(r, 1, label).font = F_BASE
        c = ws.cell(r, 2, sub(f))
        c.font, c.border, c.alignment = F_BOLD, BORDER, Alignment(horizontal="center")
        ws.cell(r, 5, sub(detail)).font = F_SUB
        s.r += 1
    c1 = s.r - 1
    status_formatting(ws, f"B{c0}:B{c1}")
    s.gap()
    ws.cell(s.r, 1, "Checks passed").font = F_BOLD
    ws.cell(s.r, 2, f'=COUNTIF(B{c0}:B{c1},"PASS")&" of "&COUNTA(B{c0}:B{c1})').font = F_BOLD
    ws.cell(s.r, 2).alignment = Alignment(horizontal="center")
    REG["n_pass_cell"] = f"Summary!$B${s.r}"

    # ===================================================================== Cover
    ws = sheets["Cover"]
    s = SW(ws, [28, 100])
    s.title(cfg.PROJECT_TITLE, "Design calculations and single line diagram for a two-transformer 33/11 kV distribution substation")
    info = [
        ("Prepared by", cfg.AUTHOR),
        ("Date", datetime.date.today().isoformat()),
        ("Status", "Learning / portfolio exercise. NOT FOR CONSTRUCTION. Inputs are assumptions, not site data."),
        ("Scope", "Load and N-1 capacity, short-circuit levels, equipment ratings, busbar thermal check, earthing grid, "
                  "overcurrent relay grading, equipment list."),
    ]
    for k, v in info:
        ws.cell(s.r, 1, k).font = F_BOLD
        ws.cell(s.r, 2, v).font = F_BASE
        s.r += 1
    s.gap()
    s.section("Sheets", 2)
    index = [
        ("Summary", "Key results and pass/fail checks"),
        ("Inputs", "All design inputs and assumptions (change these)"),
        ("Lists", "Standard ratings, conductor sizes and material constants"),
        ("Load_Growth", "Demand growth and when N-1 capacity runs out"),
        ("Short_Circuit", "Fault levels at 33 kV and 11 kV, switchgear fault ratings"),
        ("Equipment_Ratings", "Breaker, isolator and CT ratings for every circuit"),
        ("Busbar", "33 kV conductor thermal check"),
        ("Earthing", "Earth grid design: step and touch voltage"),
        ("Protection", "Overcurrent relay settings and grading curves"),
        ("BOM", "Equipment list with quantities (rates left blank)"),
        ("SLD", "Single line diagram"),
    ]
    for k, v in index:
        c = ws.cell(s.r, 1, k)
        c.hyperlink = f"#'{k}'!A1"
        c.font = Font(name=FONT, size=10, color="0563C1", underline="single")
        ws.cell(s.r, 2, v).font = F_BASE
        s.r += 1
    s.gap()
    s.section("Colour code", 2)
    legend = [("Blue text", F_INPUT, None, "Input typed by the designer"),
              ("Blue on yellow", F_INPUT, FILL_ASSUMED, "Assumption: replace with real data before relying on the result"),
              ("Green text", F_LINK, None, "Value linked from another sheet"),
              ("Black text", F_BASE, None, "Formula")]
    for k, fnt, fill, v in legend:
        c = ws.cell(s.r, 1, k)
        c.font = fnt
        if fill:
            c.fill = fill
        ws.cell(s.r, 2, v).font = F_BASE
        s.r += 1
    s.gap()
    s.section("Limits of this package", 2)
    for line in [
        "Simplified methods (per-unit fault study, IEEE 80 simplified formulas, IEC 60909 style peak factor).",
        "Numbers such as the source fault level, transformer impedance and soil resistivity are assumptions.",
        "Standard clauses, clearances, ampacity, cable sizing, DC system and civil design are not covered.",
        "A real substation design needs utility approval and a certified engineer.",
    ]:
        ws.cell(s.r, 1, "-").font = F_BASE
        ws.cell(s.r, 2, line).font = F_BASE
        s.r += 1

    return wb, dict(REG)


def main():
    root = Path(__file__).resolve().parent.parent
    sld = root / "drawings" / "SLD_33_11kV_substation.png"
    wb, _ = build(sld)
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else root / "Substation_Design_Calculations.xlsx"
    wb.save(out)
    print("Saved", out)


if __name__ == "__main__":
    sys.exit(main())
