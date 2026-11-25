# Methodology

Every calculation exists twice: once as live Excel formulas (`Substation_Design_Calculations.xlsx`) and once as plain Python (`src/calcs.py`). `src/verify_workbook.py` compares them. All methods are simplified for a learning exercise. Check the standards named below for the exact clauses, limits and tables before using any value for real work.

## 1. Load and N-1 capacity (`Load_Growth`)

- Year-0 station demand = sum of feeder peak loads x diversity factor.
- Demand in year n = year-0 demand x (1 + growth)^n, for 15 years.
- Firm capacity (N-1) = (number of transformers - 1) x rating x (1 + allowed emergency overload).
- The sheet reports the first year when demand exceeds firm capacity, and the first year when normal loading exceeds the planning limit.

## 2. Short circuit (`Short_Circuit`)

- Per-unit method on a common MVA base. Source impedance from the assumed fault level; transformer impedance from %Z.
- Balanced three-phase fault. Motor contribution and cable impedance are ignored, which is conservative.
- Two 11 kV cases: one transformer in service (coupler open) and all transformers in parallel (coupler closed). The larger value sets the 11 kV switchgear rating.
- Peak current ip = kappa x sqrt(2) x Ik'', with kappa = 1.02 + 0.98 exp(-3R/X) (IEC 60909 style).
- Switchgear short-circuit rating = next value in a standard list above Ik'' plus a margin. Rated peak withstand = 2.5 x rated short-circuit current for 50 Hz (IEC 62271 series).

## 3. Equipment ratings (`Equipment_Ratings`)

- Load current basis per circuit: transformer full-load current (33 kV and 11 kV side), both transformers for the 33 kV incomer and busbar, feeder MVA for each feeder.
- Rated current = next standard value above basis x (1 + margin). CT primary = next standard CT ratio above the same value. CT secondary is taken as 1 A.

## 4. 33 kV busbar conductor (`Busbar`)

- Minimum area from the short-circuit thermal limit: A = Ik'' x sqrt(t) / k.
- k for aluminium from the adiabatic formula in IEC 60949 using the constants listed on the `Inputs` sheet.
- Ampacity, sag, electrodynamic forces, corona and clearances are not checked.

## 5. Earthing grid (`Earthing`)

Follows the structure of IEEE Std 80 for a rectangular grid:

1. Conductor size from the fault current, fault duration and material constants.
2. Grid resistance from the Sverak formula.
3. Decrement factor from the X/R ratio and fault duration; grid current Ig = Df x Sf x 3I0.
4. Tolerable touch and step voltages with the surface-layer derating factor Cs.
5. Mesh voltage (Km, Ki, effective length) and step voltage (Ks) of the proposed grid; both must stay below the tolerable values.
6. If the ground potential rise is above the tolerable touch voltage, transferred potentials (fence, LV neutral, telecom, pipes) must be reviewed. The workbook flags this as REVIEW rather than hiding it.

The equal-area round-conductor diameter is used for flat bars (conservative). Soil is a single uniform layer; real sites need measured resistivity and a soil model.

## 6. Overcurrent grading (`Protection`)

- IEC standard-inverse curve: t = TMS x 0.14 / (PSM^0.02 - 1).
- Feeder relay is the starting point; the incomer and the 33 kV transformer relay get the smallest relay-step TMS that keeps the grading margin at the maximum 11 kV fault current.
- The 33 kV current is the 11 kV fault current scaled by the voltage ratio (three-phase fault).
- Not covered: minimum fault sensitivity, instantaneous elements, earth-fault settings, CT saturation.

## 7. Equipment list (`BOM`)

Quantities follow the single line diagram. Unit rates are left blank so the total only appears once real rates from a utility schedule of rates or a quotation are entered.

## Standards referred to

- IEC 60909 (short-circuit currents), IEC 62271 series (switchgear ratings), IEC 60949 (thermally permissible short-circuit currents), IEC 60255 (relay curves)
- IEEE Std 80 (substation grounding), IEC 60076-7 (transformer loading)
- Indian practice: CEA regulations and the utility's own technical specifications decide the final values.
