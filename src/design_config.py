"""Single source of truth for the design inputs.

Every number here is either a standard constant or an ASSUMPTION made for this
learning/portfolio exercise. Nothing is taken from a real substation. Replace the
assumed values with data from the utility / manufacturer before using the method
for anything real.

kind:
    "std"     standard / physical constant (still verify against the cited standard)
    "assumed" assumption made for this exercise - highlighted yellow in the workbook
"""

AUTHOR = "Suvadeep Mallik"
PROJECT_TITLE = "33/11 kV Substation Design Package"

# (key, label, value, unit, note, kind, number_format)
SECTIONS = [
    ("1. System", [
        ("freq", "System frequency", 50, "Hz", "Indian grid frequency", "std", "0"),
        ("V_hv", "HV nominal voltage", 33, "kV", "Sub-transmission level", "std", "0"),
        ("V_lv", "LV nominal voltage", 11, "kV", "Distribution level", "std", "0"),
        ("S_src", "Source fault level at 33 kV bus", 750, "MVA",
         "ASSUMED. Obtain the real value from the utility / grid substation.", "assumed", "#,##0"),
        ("XR_src", "Source X/R ratio", 10, "-", "ASSUMED typical order of magnitude.", "assumed", "0.0"),
        ("S_base", "Calculation base power", 100, "MVA", "Any base gives the same answer", "std", "0"),
    ]),
    ("2. Power transformer", [
        ("n_tx", "Number of transformers", 2, "nos", "Two-transformer station (N-1 capable)", "std", "0"),
        ("S_tx", "Rating per transformer", 10, "MVA", "ONAN, 33/11 kV, Dyn11", "std", "0.0"),
        ("Z_tx", "Percentage impedance", 0.08, "%",
         "ASSUMED. Replace with guaranteed %Z from the transformer specification.", "assumed", "0.00%"),
        ("XR_tx", "Transformer X/R ratio", 15, "-", "ASSUMED typical order of magnitude.", "assumed", "0.0"),
        ("ov_n1", "Emergency overload allowed in N-1", 0.0, "%",
         "0 % = no overload credit (conservative). Check IEC 60076-7 / utility policy.", "assumed", "0%"),
        ("lim_norm", "Normal-operation loading limit", 0.80, "% of rating",
         "ASSUMED planning limit. Replace with the utility planning criterion.", "assumed", "0%"),
    ]),
    ("3. Load", [
        ("pf", "Power factor", 0.90, "-", "ASSUMED", "assumed", "0.00"),
        ("growth", "Load growth", 0.05, "% / year", "ASSUMED", "assumed", "0.0%"),
        ("diversity", "Diversity factor on feeder peaks", 0.80, "-",
         "ASSUMED. Not all feeders peak together.", "assumed", "0.00"),
        ("sat_kva", "Station auxiliary transformer rating", 100, "kVA", "ASSUMED", "assumed", "0"),
    ]),
    ("4. Switchgear", [
        ("t_k", "Rated short-time withstand duration", 3, "s", "ASSUMED. Confirm with utility specification.", "assumed", "0"),
        ("m_isc", "Margin on calculated fault current", 0.10, "%", "ASSUMED headroom for network growth", "assumed", "0%"),
        ("m_i", "Margin on rated current", 0.10, "%", "ASSUMED", "assumed", "0%"),
    ]),
    ("5. Busbar conductor (33 kV outdoor)", [
        ("t_bus", "Fault duration for thermal check", 1.0, "s", "ASSUMED back-up clearing time", "assumed", "0.0"),
        ("th_i", "Initial conductor temperature", 40, "deg C", "ASSUMED", "assumed", "0"),
        ("th_f", "Final (short-circuit) temperature", 200, "deg C",
         "ASSUMED. Confirm limit for the conductor type (IEC 60865-1).", "assumed", "0"),
        ("Qc", "Volumetric heat capacity of aluminium", 2.5e6, "J/(K.m3)", "IEC 60949 constant (verify)", "std", "0.00E+00"),
        ("beta", "Reciprocal temp. coefficient of resistance (Al)", 228, "K", "IEC 60949 constant (verify)", "std", "0"),
        ("rho20", "Resistivity of aluminium at 20 deg C", 28.264e-9, "ohm.m", "IEC 60949 constant (verify)", "std", "0.000E+00"),
    ]),
    ("6. Earthing grid (IEEE 80 method)", [
        ("rho_soil", "Soil resistivity", 100, "ohm.m", "ASSUMED. Must be measured on site (Wenner 4-pin).", "assumed", "0"),
        ("rho_s", "Surface layer resistivity (crushed rock)", 3000, "ohm.m", "ASSUMED", "assumed", "0"),
        ("hs", "Surface layer thickness", 0.10, "m", "ASSUMED", "assumed", "0.00"),
        ("body", "Body weight for tolerable voltage", 70, "kg", "Use 50 or 70", "std", "0"),
        ("t_f", "Fault duration (shock and decrement factor)", 0.5, "s", "ASSUMED protection clearing time", "assumed", "0.0"),
        ("t_c", "Fault duration for conductor sizing", 1.0, "s", "ASSUMED back-up clearing time", "assumed", "0.0"),
        ("I_3I0", "Maximum ground fault current (3I0)", 8.0, "kA",
         "ASSUMED. Depends on 33 kV source neutral earthing; take from the network study.", "assumed", "0.0"),
        ("Sf", "Current division factor", 1.0, "-", "1.0 = all fault current returns through the grid (conservative)", "assumed", "0.00"),
        ("XR_g", "X/R at ground fault location", 10, "-", "ASSUMED", "assumed", "0.0"),
        ("Lx", "Grid length", 48, "m", "ASSUMED yard size (keep a multiple of spacing)", "assumed", "0.0"),
        ("Ly", "Grid width", 36, "m", "ASSUMED yard size (keep a multiple of spacing)", "assumed", "0.0"),
        ("D", "Conductor spacing", 3, "m", "Design variable (4 m gives almost no margin on touch voltage)", "assumed", "0.0"),
        ("h", "Burial depth", 0.5, "m", "ASSUMED", "assumed", "0.00"),
        ("rods_n", "Number of earth rods", 24, "nos", "Perimeter and corner rods", "assumed", "0"),
        ("rods_L", "Length of each rod", 3.0, "m", "ASSUMED", "assumed", "0.0"),
        ("rods_perim", "Rods on perimeter and corners? (1 = yes)", 1, "-", "Affects Kii in the mesh voltage formula", "std", "0"),
        ("mat", "Grid conductor material", "Zinc-coated steel", "-", "Pick from list (GI flat is common in India)", "std", "@"),
        ("min_area", "Minimum conductor area (mechanical/corrosion)", 300, "mm2",
         "ASSUMED. Replace with utility specification.", "assumed", "0"),
        ("T_amb", "Ambient temperature for conductor sizing", 40, "deg C", "ASSUMED", "assumed", "0"),
    ]),
    ("7. Overcurrent protection", [
        ("pk_f", "Feeder relay pickup (x full-load current)", 1.25, "-", "ASSUMED", "assumed", "0.00"),
        ("pk_inc", "11 kV incomer relay pickup (x FLC)", 1.25, "-", "ASSUMED", "assumed", "0.00"),
        ("pk_hv", "33 kV transformer relay pickup (x FLC)", 1.25, "-", "ASSUMED", "assumed", "0.00"),
        ("tms_f", "Feeder relay time multiplier setting", 0.10, "-", "ASSUMED starting point", "assumed", "0.000"),
        ("margin", "Grading margin between relays", 0.30, "s", "ASSUMED", "assumed", "0.00"),
        ("tms_step", "Relay TMS step", 0.025, "-", "ASSUMED relay setting resolution", "assumed", "0.000"),
        ("k_si", "IDMT curve constant k (standard inverse)", 0.14, "-", "IEC 60255 standard inverse", "std", "0.00"),
        ("a_si", "IDMT curve exponent alpha", 0.02, "-", "IEC 60255 standard inverse", "std", "0.00"),
    ]),
]

# Outgoing 11 kV feeders: (name, peak load in MVA). ASSUMED illustrative loads.
FEEDERS = [
    ("F1 Urban residential", 2.0),
    ("F2 Commercial", 1.5),
    ("F3 Agricultural", 1.8),
    ("F4 Industrial", 2.2),
    ("F5 Rural", 1.2),
    ("F6 Semi-urban mixed", 1.6),
]

# Standard value lists (ascending). IEC R10-type preferred values; confirm with the utility spec.
STD_ISC_KA = [6.3, 8, 10, 12.5, 16, 20, 25, 31.5, 40, 50, 63]
STD_RATED_A = [400, 630, 800, 1250, 1600, 2000, 2500, 3150, 4000]
STD_CT_PRIMARY_A = [50, 75, 100, 150, 200, 300, 400, 600, 800, 1000, 1200, 1500, 2000]
# ACSR conductors by nominal aluminium area (mm2)
ACSR = [("Dog", 100), ("Wolf", 150), ("Panther", 200), ("Zebra", 400)]
# Flat bars for the earthing grid: (name, area mm2)
FLATS = [("25 x 3", 75), ("25 x 6", 150), ("40 x 6", 240), ("50 x 6", 300),
         ("50 x 8", 400), ("50 x 10", 500), ("75 x 10", 750)]
# IEEE 80-2013 Table 2 style material constants: (name, alpha_r, K0, Tm, rho_r, TCAP)
EARTH_MATERIALS = [
    ("Copper, hard-drawn", 0.00381, 242, 1084, 1.78, 3.42),
    ("Steel (mild, 1020)", 0.0016, 605, 1510, 15.9, 3.28),
    ("Zinc-coated steel", 0.0032, 293, 419, 20.1, 3.93),
]

N_YEARS = 15  # planning horizon in Load_Growth


def defaults():
    """Flat dict of input values by key (feeder loads under f1..fN)."""
    d = {}
    for _, items in SECTIONS:
        for key, _label, value, *_ in items:
            d[key] = value
    for i, (_n, mva) in enumerate(FEEDERS, start=1):
        d[f"f{i}"] = mva
    return d
