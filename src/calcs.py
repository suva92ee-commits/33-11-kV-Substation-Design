"""Pure-Python implementation of every design calculation.

This is deliberately an INDEPENDENT second implementation of what the Excel
workbook does with formulas. `verify_workbook.py` compares the two, so a typo in
either one shows up as a mismatch. The SLD labels are also taken from here.

Methods used (all simplified, for learning purposes):
  - Short circuit: per-unit impedance method, balanced three-phase fault,
    no motor contribution, no cable impedance (conservative).
  - Peak current: IEC 60909 style kappa factor.
  - Earthing: IEEE 80 simplified formulas (rectangular grid, Sverak resistance).
  - Overcurrent grading: IEC standard-inverse curve.
"""
import math

import design_config as cfg


def next_std(values, x):
    """Smallest value in the ascending list that is >= x, else None."""
    for v in values:
        if v >= x:
            return v
    return None


def ceil_step(x, step):
    return math.ceil(x / step - 1e-9) * step


def compute(inp=None):
    p = dict(cfg.defaults())
    if inp:
        p.update(inp)
    r = {}
    sqrt3 = math.sqrt(3)

    # ---------------- Load and transformer capacity ----------------
    feeders = [p[f"f{i}"] for i in range(1, len(cfg.FEEDERS) + 1)]
    r["sum_feeders"] = sum(feeders)
    r["D0"] = r["sum_feeders"] * p["diversity"]
    r["D0_MW"] = r["D0"] * p["pf"]
    r["cap_inst"] = p["n_tx"] * p["S_tx"]
    r["cap_n1"] = (p["n_tx"] - 1) * p["S_tx"] * (1 + p["ov_n1"])
    r["load_norm0"] = r["D0"] / r["cap_inst"]
    r["load_n10"] = r["D0"] / r["cap_n1"]
    r["breach_year"] = ">15"
    r["watch_year"] = ">15"
    for n in range(0, cfg.N_YEARS + 1):
        d = r["D0"] * (1 + p["growth"]) ** n
        if r["breach_year"] == ">15" and d > r["cap_n1"]:
            r["breach_year"] = n
        if r["watch_year"] == ">15" and d / r["cap_inst"] > p["lim_norm"]:
            r["watch_year"] = n

    # ---------------- Short circuit ----------------
    sb = p["S_base"]

    def split(z, xr):
        rr = z / math.sqrt(1 + xr ** 2)
        return rr, rr * xr

    r["Zs"] = sb / p["S_src"]
    r["Rs"], r["Xs"] = split(r["Zs"], p["XR_src"])
    r["Zt"] = p["Z_tx"] * sb / p["S_tx"]
    r["Rt"], r["Xt"] = split(r["Zt"], p["XR_tx"])

    def fault(n_parallel, vkv):
        rr = r["Rs"] + r["Rt"] / n_parallel
        xx = r["Xs"] + r["Xt"] / n_parallel
        z = math.hypot(rr, xx)
        ik = sb / (sqrt3 * vkv * z)
        kappa = 1.02 + 0.98 * math.exp(-3 * rr / xx)
        return z, ik, sb / z, xx / rr, kappa, kappa * math.sqrt(2) * ik

    (r["Z_a"], r["Ik_a"], r["S_a"], r["XR_a"], r["k_a"], r["ip_a"]) = fault(1, p["V_lv"])
    (r["Z_b"], r["Ik_b"], r["S_b"], r["XR_b"], r["k_b"], r["ip_b"]) = fault(p["n_tx"], p["V_lv"])
    r["Ik33"] = p["S_src"] / (sqrt3 * p["V_hv"])
    r["k33"] = 1.02 + 0.98 * math.exp(-3 / p["XR_src"])
    r["ip33"] = r["k33"] * math.sqrt(2) * r["Ik33"]
    r["Ik11"] = max(r["Ik_a"], r["Ik_b"])
    r["Isc33_rated"] = next_std(cfg.STD_ISC_KA, r["Ik33"] * (1 + p["m_isc"]))
    r["Isc11_rated"] = next_std(cfg.STD_ISC_KA, r["Ik11"] * (1 + p["m_isc"]))
    r["ip33_rated"] = 2.5 * r["Isc33_rated"]
    r["ip11_rated"] = 2.5 * r["Isc11_rated"]

    # ---------------- Equipment ratings ----------------
    r["I_hv_flc"] = p["S_tx"] * 1000 / (sqrt3 * p["V_hv"])
    r["I_lv_flc"] = p["S_tx"] * 1000 / (sqrt3 * p["V_lv"])
    basis = {
        "inc33": p["n_tx"] * r["I_hv_flc"],
        "cpl33": r["I_hv_flc"],
        "txhv": r["I_hv_flc"],
        "bus33": p["n_tx"] * r["I_hv_flc"],
        "inc11": r["I_lv_flc"],
        "cpl11": r["I_lv_flc"],
        "bus11": r["I_lv_flc"],
    }
    for i, mva in enumerate(feeders, start=1):
        basis[f"fdr{i}"] = mva * 1000 / (sqrt3 * p["V_lv"])
    r["circuits"] = list(basis)
    for c, b in basis.items():
        design = b * (1 + p["m_i"])
        r[f"{c}_basis"] = b
        r[f"{c}_design"] = design
        r[f"{c}_cb"] = next_std(cfg.STD_RATED_A, design)
        r[f"{c}_ct"] = next_std(cfg.STD_CT_PRIMARY_A, design)

    # ---------------- Busbar thermal check (33 kV conductor) ----------------
    r["k_al"] = math.sqrt(p["Qc"] * (p["beta"] + 20) / p["rho20"]
                          * math.log((p["beta"] + p["th_f"]) / (p["beta"] + p["th_i"]))) * 1e-6
    r["A_min_bus"] = r["Ik33"] * 1000 * math.sqrt(p["t_bus"]) / r["k_al"]
    areas = [a for _, a in cfg.ACSR]
    r["A_sel_bus"] = next_std(areas, r["A_min_bus"])
    r["bus_name"] = dict((a, n) for n, a in cfg.ACSR).get(r["A_sel_bus"], "none")
    r["J_bus"] = r["bus33_design"] / r["A_sel_bus"] if r["A_sel_bus"] else None

    # ---------------- Earthing grid ----------------
    mat = {m[0]: m for m in cfg.EARTH_MATERIALS}[p["mat"]]
    _, ar, k0, tm, rr_, tcap = mat
    r["A_req_earth"] = p["I_3I0"] / math.sqrt(
        (tcap * 1e-4 / (p["t_c"] * ar * rr_)) * math.log(1 + (tm - p["T_amb"]) / (k0 + p["T_amb"])))
    flats = [a for _, a in cfg.FLATS]
    r["A_sel_earth"] = next_std(flats, max(r["A_req_earth"], p["min_area"]))
    r["flat_name"] = dict((a, n) for n, a in cfg.FLATS)[r["A_sel_earth"]]
    r["d_cond"] = 2 * math.sqrt(r["A_sel_earth"] / math.pi) / 1000
    r["nA"] = math.floor(p["Ly"] / p["D"] + 0.5) + 1  # same rounding as Excel ROUND()
    r["nB"] = math.floor(p["Lx"] / p["D"] + 0.5) + 1
    r["LC"] = r["nA"] * p["Lx"] + r["nB"] * p["Ly"]
    r["Agrid"] = p["Lx"] * p["Ly"]
    r["LP"] = 2 * (p["Lx"] + p["Ly"])
    r["LR"] = p["rods_n"] * p["rods_L"]
    r["LT"] = r["LC"] + r["LR"]
    A, h = r["Agrid"], p["h"]
    r["Rg"] = p["rho_soil"] * (1 / r["LT"] + 1 / math.sqrt(20 * A) * (1 + 1 / (1 + h * math.sqrt(20 / A))))
    r["Ta_dc"] = p["XR_g"] / (2 * math.pi * p["freq"])
    r["Df"] = math.sqrt(1 + (r["Ta_dc"] / p["t_f"]) * (1 - math.exp(-2 * p["t_f"] / r["Ta_dc"])))
    r["Ig"] = r["Df"] * p["Sf"] * p["I_3I0"] * 1000
    r["GPR"] = r["Ig"] * r["Rg"]
    r["Cs"] = 1 - 0.09 * (1 - p["rho_soil"] / p["rho_s"]) / (2 * p["hs"] + 0.09)
    r["Kb"] = 0.157 if p["body"] == 70 else 0.116
    r["Etouch"] = (1000 + 1.5 * r["Cs"] * p["rho_s"]) * r["Kb"] / math.sqrt(p["t_f"])
    r["Estep"] = (1000 + 6 * r["Cs"] * p["rho_s"]) * r["Kb"] / math.sqrt(p["t_f"])
    r["na"] = 2 * r["LC"] / r["LP"]
    r["nb"] = math.sqrt(r["LP"] / (4 * math.sqrt(A)))
    r["n_eff"] = r["na"] * r["nb"]
    n = r["n_eff"]
    r["Ki"] = 0.644 + 0.148 * n
    r["Kh"] = math.sqrt(1 + h / 1.0)
    r["Kii"] = 1.0 if p["rods_perim"] == 1 else 1 / (2 * n) ** (2 / n)
    D, d = p["D"], r["d_cond"]
    r["Km"] = (1 / (2 * math.pi)) * (
        math.log(D ** 2 / (16 * h * d) + (D + 2 * h) ** 2 / (8 * D * d) - h / (4 * d))
        + (r["Kii"] / r["Kh"]) * math.log(8 / (math.pi * (2 * n - 1))))
    r["LM"] = r["LC"] + (1.55 + 1.22 * (p["rods_L"] / math.hypot(p["Lx"], p["Ly"]))) * r["LR"]
    r["Em"] = p["rho_soil"] * r["Km"] * r["Ki"] * r["Ig"] / r["LM"]
    r["Ks"] = (1 / math.pi) * (1 / (2 * h) + 1 / (D + h) + (1 / D) * (1 - 0.5 ** (n - 2)))
    r["LS"] = 0.75 * r["LC"] + 0.85 * r["LR"]
    r["Es"] = p["rho_soil"] * r["Ks"] * r["Ki"] * r["Ig"] / r["LS"]

    # ---------------- Overcurrent grading ----------------
    k, a = p["k_si"], p["a_si"]

    def t_op(tms, psm):
        return tms * k / (psm ** a - 1)

    r["If_prot"] = r["Ik_a"] * 1000  # worst-case: one transformer in service
    r["I_fdr_max"] = max(r[f"fdr{i}_basis"] for i in range(1, len(feeders) + 1))
    r["Ip_f"] = p["pk_f"] * r["I_fdr_max"]
    r["PSM_f"] = r["If_prot"] / r["Ip_f"]
    r["t_f_relay"] = t_op(p["tms_f"], r["PSM_f"])
    r["Ip_inc"] = p["pk_inc"] * r["I_lv_flc"]
    r["PSM_inc"] = r["If_prot"] / r["Ip_inc"]
    r["TMS_inc_req"] = (r["t_f_relay"] + p["margin"]) * (r["PSM_inc"] ** a - 1) / k
    r["TMS_inc"] = ceil_step(r["TMS_inc_req"], p["tms_step"])
    r["t_inc"] = t_op(r["TMS_inc"], r["PSM_inc"])
    r["I_hv_fault"] = r["If_prot"] * p["V_lv"] / p["V_hv"]
    r["Ip_hv"] = p["pk_hv"] * r["I_hv_flc"]
    r["PSM_hv"] = r["I_hv_fault"] / r["Ip_hv"]
    r["TMS_hv_req"] = (r["t_inc"] + p["margin"]) * (r["PSM_hv"] ** a - 1) / k
    r["TMS_hv"] = ceil_step(r["TMS_hv_req"], p["tms_step"])
    r["t_hv"] = t_op(r["TMS_hv"], r["PSM_hv"])
    r["margin_inc"] = r["t_inc"] - r["t_f_relay"]
    r["margin_hv"] = r["t_hv"] - r["t_inc"]
    return r


if __name__ == "__main__":
    res = compute()
    for key in ["D0", "load_n10", "breach_year", "Ik33", "Ik_a", "Ik_b", "Ik11", "Isc33_rated", "Isc11_rated",
                "inc33_cb", "inc11_cb", "k_al", "A_min_bus", "bus_name", "A_req_earth", "flat_name", "Rg", "Ig",
                "GPR", "Etouch", "Estep", "Em", "Es", "Km", "Ki", "t_f_relay", "TMS_inc", "t_inc", "TMS_hv", "t_hv"]:
        v = res[key]
        print(f"{key:14s} {v:.4f}" if isinstance(v, float) else f"{key:14s} {v}")
