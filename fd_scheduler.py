import os
import pandas as pd
import gurobipy as gp
from gurobipy import GRB
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# =====================================================================
# SETTINGS  (change these, nothing else needs touching)
# =====================================================================
INPUT_FILE   = "FD_availability_template.xlsx"   # one sheet per person
OUTPUT_FILE  = "fd_schedule_output.xlsx"
SKIP_SHEETS  = {"Instructions", "Blank Template"}

DAYS         = ["Mon", "Tue", "Wed", "Thu", "Fri"]
SLOTS_PER_DAY = {"Mon": 16, "Tue": 16, "Wed": 16, "Thu": 16, "Fri": 12}   # 30-min slots, 9:00 start
SLOT_HRS     = 0.5
MIN_SHIFT    = 4      # slots  -> 2 hours
MAX_SHIFT    = 10     # slots  -> 5 hours
MIN_WEEK_HRS = 2.0    # every RETURNING person must get at least one real shift
MIN_HIRE_HRS = 4.0    # a hired APPLICANT must get at least this many hrs/wk (or their max, if lower)

# Objective weights (priority order: gap > fair > sum > over > pref)
W_GAP, W_FAIR, W_SUM, W_OVER, W_PREF = 1000, 150, 15, 2, 1


# =====================================================================
# 1. LOAD DATA
# =====================================================================
def slot_label(k):
    m = 9 * 60 + 30 * k
    h, mm = divmod(m, 60)
    return f"{h if h <= 12 else h - 12}:{mm:02d}{'a' if h < 12 else 'p'}"

def is_red(cell):
    f = cell.fill
    if f is None or f.fill_type != "solid":
        return False
    rgb = f.fgColor.rgb
    if not isinstance(rgb, str) or len(rgb) != 8:
        return False
    r, g, b = int(rgb[2:4], 16), int(rgb[4:6], 16), int(rgb[6:8], 16)
    return r > 180 and g < 120 and b < 120

def load_data(path):
    # Build the master slot list: slots[t] = (day, index_in_day, label)
    slots, day_slots = [], {d: [] for d in DAYS}
    for d in DAYS:
        for k in range(SLOTS_PER_DAY[d]):
            day_slots[d].append(len(slots))
            slots.append((d, k, slot_label(k)))

    wb = load_workbook(path)
    staff, ideal, mx, avail, pref = [], {}, {}, {}, {}
    for ws in wb.worksheets:
        if ws.title in SKIP_SHEETS:
            continue
        name = str(ws["B1"].value).strip()
        ideal[name], mx[name] = float(ws["B2"].value), float(ws["B3"].value)
        staff.append(name)
        for t, (d, k, _) in enumerate(slots):
            cell = ws.cell(row=6 + k, column=2 + DAYS.index(d))   # grid starts at B6
            txt = str(cell.value).strip().lower() if cell.value is not None else ""
            ok = (txt in ("", "a", "p")) and not is_red(cell)
            avail[name, t] = 1 if ok else 0
            pref[name, t] = 1 if (txt == "p" and ok) else 0
    return slots, day_slots, staff, ideal, mx, avail, pref



# =====================================================================
# 1b. INPUT CHECKS  (runs BEFORE the solver; stops the run if something is wrong)
# =====================================================================
class StopRun(Exception):
    """Raised to stop a notebook cell cleanly after problems have been printed."""
    def _render_traceback_(self):
        return ["Run stopped: fix the problems listed above in the Excel file(s), then re-run."]

KNOWN_TEXT = {"", "a", "p", "class", "lab", "closed"}   # anything else (uncoloured) is "can't read"

def _num(v):
    try:
        x = float(v)
        return x if x == x else None
    except (TypeError, ValueError):
        return None

def check_inputs(files):
    """
    files = [(path, "returning"), (path, "applicant"), ...]
    Prints PROBLEMS (stop the run) and HEADS-UPs (do not stop).
    Returns the set of applicant names that can't be hired (so the hiring step can skip them).
    """
    hard, soft, cannot_hire = [], [], set()
    seen, ret_max_total, desk_hrs = {}, 0.0, sum(SLOTS_PER_DAY.values()) * SLOT_HRS

    for path, role in files:
        if not os.path.exists(path):
            hard.append(f"File not found: {path}")
            continue
        sheets = [ws for ws in load_workbook(path).worksheets if ws.title not in SKIP_SHEETS]
        if not sheets:
            hard.append(f"{path}: no person sheets found (only {sorted(SKIP_SHEETS)} or nothing at all)")
            continue

        for ws in sheets:
            tab = ws.title
            name = str(ws["B1"].value).strip() if ws["B1"].value not in (None, "") else ""
            if name.lower() in ("", "nan", "none"):
                hard.append(f"[{tab}] name cell B1 is empty")
                name = tab
            elif name != tab:
                soft.append(f"[{tab}] tab is called '{tab}' but B1 says '{name}' (the code uses B1)")
            if name in seen:
                hard.append(f"Duplicate name '{name}': appears in {seen[name]} and in {os.path.basename(path)} / {tab}")
            seen[name] = f"{os.path.basename(path)} / {tab}"

            # ---- ideal / max ----
            ideal_v, max_v = _num(ws["B2"].value), _num(ws["B3"].value)
            if ideal_v is None:
                hard.append(f"[{name}] ideal hrs (B2) is missing or not a number: '{ws['B2'].value}'")
            if max_v is None:
                hard.append(f"[{name}] max hrs (B3) is missing or not a number: '{ws['B3'].value}'")
            if ideal_v is not None and max_v is not None:
                if ideal_v < 0 or max_v < 0:
                    hard.append(f"[{name}] ideal/max hours can't be negative ({ideal_v}, {max_v})")
                elif ideal_v > max_v:
                    hard.append(f"[{name}] ideal ({ideal_v:g}) is above max ({max_v:g})")

            # ---- read the grid ----
            free, bad, red_p = {}, [], []
            for d in DAYS:
                free[d] = []
                for k in range(SLOTS_PER_DAY[d]):
                    c = ws.cell(row=6 + k, column=2 + DAYS.index(d))
                    txt = str(c.value).strip().lower() if c.value is not None else ""
                    red = is_red(c)
                    if txt not in KNOWN_TEXT and not red:
                        bad.append(f"{d} {slot_label(k)} = '{c.value}'")
                    if txt == "p" and red:
                        red_p.append(f"{d} {slot_label(k)}")
                    free[d].append(txt in ("", "a", "p") and not red)
            if bad:
                shown = "; ".join(bad[:5]) + (f"; ...and {len(bad) - 5} more" if len(bad) > 5 else "")
                hard.append(f"[{name}] can't read {len(bad)} cell(s): {shown}. "
                            f"Use P, A, Class, or leave blank (or fill the cell red if blocked)")
            if red_p:
                soft.append(f"[{name}] marked P but filled red (treated as unavailable): {', '.join(red_p[:4])}")

            # ---- can this person work a legal shift at all? ----
            best, longest, longest_day = 0.0, 0, None
            for d in DAYS:
                run = dmax = 0
                for ok in free[d]:
                    run = run + 1 if ok else 0
                    dmax = max(dmax, run)
                if dmax >= MIN_SHIFT:
                    best += min(dmax, MAX_SHIFT) * SLOT_HRS     # best case: one block per day
                if dmax > longest:
                    longest, longest_day = dmax, d
            need_hrs = MIN_SHIFT * SLOT_HRS
            why_none = (f"no free stretch of {need_hrs:g}+ hrs in a row on any day "
                        f"(longest free stretch: {longest * SLOT_HRS:g} hr{'s' if longest * SLOT_HRS != 1 else ''}"
                        f"{' on ' + longest_day if longest_day else ''})")

            if role == "returning":
                if max_v is not None:
                    ret_max_total += max_v
                floor = MIN_WEEK_HRS
                if best < floor:
                    hard.append(f"[{name}] {why_none}, but every returning person must get at least "
                                f"{floor:g} hrs. Fix their availability, or remove them from this semester's file")
                elif max_v is not None and max_v < floor:
                    hard.append(f"[{name}] max hrs ({max_v:g}) is below the {floor:g}-hr minimum every returning person gets")
                elif ideal_v is not None and best < ideal_v:
                    soft.append(f"[{name}] free time only allows up to {best:g} hrs/wk (in {need_hrs:g}-{MAX_SHIFT*SLOT_HRS:g} hr "
                                f"blocks, one per day), below their ideal of {ideal_v:g}")
            else:  # applicant
                floor = min(MIN_HIRE_HRS, max_v) if max_v is not None else MIN_HIRE_HRS
                if best < need_hrs:
                    cannot_hire.add(name)
                    soft.append(f"[{name}] (applicant) can't be hired: {why_none}. Left out of the hiring comparison")
                elif best < floor or (max_v is not None and max_v < need_hrs):
                    cannot_hire.add(name)
                    soft.append(f"[{name}] (applicant) can't be hired: free time only allows up to {best:g} hrs/wk, "
                                f"below the {floor:g}-hr minimum for a hire. Left out of the hiring comparison")

    if ret_max_total and ret_max_total < desk_hrs:
        soft.append(f"Returning staff's max hours add up to {ret_max_total:g}, below the {desk_hrs:g} desk hours, "
                    f"so some gaps are unavoidable")

    print("=" * 50 + "\n   INPUT CHECK\n" + "=" * 50)
    if hard:
        print(f"\nPROBLEMS ({len(hard)}) - the run was stopped. Fix these in the Excel file(s), then re-run:")
        for i, msg in enumerate(hard, 1):
            print(f"  {i}. {msg}")
    if soft:
        print(f"\nHEADS-UP ({len(soft)}) - not blocking:")
        for msg in soft:
            print(f"  - {msg}")
    if hard:
        raise StopRun
    print("\nAll checks passed.\n")
    return cannot_hire


def explain_conflict(m):
    """Turn Gurobi's conflicting-constraint list into plain sentences."""
    m.computeIIS()
    people, general = {}, []
    for c in m.getConstrs():
        if not c.IISConstr:
            continue
        p = c.ConstrName.split("|")
        kind = p[0]
        slot_txt = lambda t: f"{slots[int(t)][0]} {slots[int(t)][2]}"
        if kind == "cover":
            general.append(f"coverage of {slot_txt(p[1])}")
        elif kind == "max_hours":
            people.setdefault(p[1], set()).add("weekly max-hours cap")
        elif kind == "everyone_works":
            people.setdefault(p[1], set()).add("minimum weekly hours")
        elif kind in ("start_def", "min_shift"):
            people.setdefault(p[1], set()).add(f"{MIN_SHIFT*SLOT_HRS:g}-hr minimum shift ({slot_txt(p[2])})")
        elif kind == "one_shift":
            people.setdefault(p[1], set()).add(f"one shift per day ({p[2]})")
        elif kind == "max_len":
            people.setdefault(p[1], set()).add(f"{MAX_SHIFT*SLOT_HRS:g}-hr shift cap ({p[2]})")
        elif kind in ("dev", "fair"):
            people.setdefault(p[1], set()).add("ideal-hours tracking")
    lines = ["These rules cannot all be met at the same time:"]
    for who in sorted(people):
        lines.append(f"  - {who}: " + "; ".join(sorted(people[who])))
    lines += [f"  - {g}" for g in general]
    lines.append("Check the availability, ideal and max for the people listed above.")
    return "\n".join(lines)

# =====================================================================
# 2. BUILD + SOLVE MODEL
# =====================================================================
def solve(slots, day_slots, staff, ideal, mx, avail, pref, time_limit=120,
          min_hrs=None, gap_tol=0.01, quiet=False):
    S, T = staff, range(len(slots))
    min_hrs = min_hrs or {s: MIN_WEEK_HRS for s in S}   # weekly floor per person
    m = gp.Model("FD_Schedule")
    if quiet:
        m.Params.OutputFlag = 0

    # ---- Decision variables ----
    # x / start only exist where the person is available (never-when-unavailable is
    # enforced by simply not creating the variable). xv() returns 0 for missing ones.
    x, start = {}, {}
    for s in S:
        for t in T:
            if avail[s, t]:
                x[s, t]     = m.addVar(vtype=GRB.BINARY, name=f"x_{s}_{t}")      # person works slot
                start[s, t] = m.addVar(vtype=GRB.BINARY, name=f"start_{s}_{t}")  # shift begins at slot
    xv = lambda s, t: x.get((s, t), 0)

    gap   = m.addVars(T, vtype=GRB.BINARY, name="gap")        # slot uncovered
    under = m.addVars(S, lb=0, name="under")                  # hrs below ideal
    over  = m.addVars(S, lb=0, name="over")                   # hrs above ideal
    maxshort = m.addVar(lb=0, name="maxshort")                # worst-off person's shortfall (fraction)

    hrs = {s: SLOT_HRS * gp.quicksum(xv(s, t) for t in T) for s in S}
    idl = {s: max(ideal[s], 1.0) for s in S}                  # guard against ideal = 0

    # ---- Hard constraints ----
    for t in T:                                               # one person per desk (or flagged gap)
        m.addConstr(gp.quicksum(xv(s, t) for s in S) + gap[t] == 1, name=f"cover|{t}")

    for s in S:
        m.addConstr(hrs[s] <= mx[s], name=f"max_hours|{s}")         # never over max
        m.addConstr(hrs[s] >= min_hrs[s], name=f"everyone_works|{s}")

        for d in DAYS:
            ts = day_slots[d]
            n = len(ts)
            for k, t in enumerate(ts):
                if (s, t) not in x:
                    continue
                prev = xv(s, ts[k - 1]) if k > 0 else 0
                m.addConstr(start[s, t] >= x[s, t] - prev, name=f"start_def|{s}|{t}")    # start definition
                m.addConstr(start[s, t] <= x[s, t], name=f"start_def|{s}|{t}")
                if k > 0:
                    m.addConstr(start[s, t] <= 1 - prev, name=f"start_def|{s}|{t}")
                if k + MIN_SHIFT > n:                         # not enough room left -> can't start
                    m.addConstr(start[s, t] == 0, name=f"min_shift|{s}|{t}")
                else:                                         # min 2-hr block
                    for j in range(1, MIN_SHIFT):
                        m.addConstr(start[s, t] <= xv(s, ts[k + j]), name=f"min_shift|{s}|{t}")
            m.addConstr(gp.quicksum(start[s, t] for t in ts if (s, t) in start) <= 1, name=f"one_shift|{s}|{d}")  # one shift/day
            m.addConstr(gp.quicksum(xv(s, t) for t in ts) <= MAX_SHIFT, name=f"max_len|{s}|{d}")                # max 5 hrs

        # deviation from ideal
        m.addConstr(hrs[s] - ideal[s] == over[s] - under[s], name=f"dev|{s}")
        m.addConstr(under[s] / idl[s] <= maxshort, name=f"fair|{s}")            # min-max fairness

    # ---- Objective ----
    obj = (W_GAP  * gp.quicksum(gap[t] for t in T)
         + W_FAIR * maxshort
         + W_SUM  * gp.quicksum(under[s] / idl[s] for s in S)
         + W_OVER * gp.quicksum(over[s] / idl[s] for s in S)
         - W_PREF * gp.quicksum(pref[s, t] * x[s, t] for (s, t) in x))
    m.setObjective(obj, GRB.MINIMIZE)

    m.Params.TimeLimit = time_limit
    m.Params.MIPGap = gap_tol
    m.optimize()

    if m.SolCount == 0:
        if m.Status == GRB.INFEASIBLE:
            solve.last_reason = explain_conflict(m)
        else:
            solve.last_reason = f"no solution found (Gurobi status {m.Status}); try a longer time_limit"
        if not quiet:
            print("\nINFEASIBLE / NO SOLUTION\n" + solve.last_reason)
        return None

    assign = {t: next((s for s in S if (s, t) in x and x[s, t].X > 0.5), None) for t in T}
    return assign, {s: SLOT_HRS * sum(1 for t in T if assign[t] == s) for s in S}


# =====================================================================
# 3. REPORT RESULTS
# =====================================================================
def report(assign, hours, slots, staff, ideal, mx, pref, out_file):
    # ---- hours summary ----
    rows = []
    for s in staff:
        pref_hrs = SLOT_HRS * sum(pref[s, t] for t in assign if assign[t] == s)
        rows.append({"Staff": s, "Assigned": hours[s], "Ideal": ideal[s], "Max": mx[s],
                     "vs Ideal": hours[s] - ideal[s], "Preferred hrs": pref_hrs})
    summary = pd.DataFrame(rows)
    print("\n--- HOURS SUMMARY ---")
    print(summary.to_string(index=False))

    # ---- gaps (where you need to hire) ----
    gaps = [(slots[t][0], slots[t][2]) for t in assign if assign[t] is None]
    print("\n--- UNCOVERED SLOTS (nobody available) ---")
    print(gaps if gaps else "None - full coverage")

    # ---- weekly grid ----
    grid = pd.DataFrame(index=[slot_label(k) for k in range(16)], columns=DAYS, dtype=object)
    for t, s in assign.items():
        d, k, lab = slots[t]
        grid.loc[lab, d] = s
    grid = grid.fillna("")
    print("\n--- WEEKLY SCHEDULE ---")
    print(grid.to_string())

    # ---- shift list ----
    print("\n--- SHIFTS ---")
    for d in DAYS:
        ts = [t for t in assign if slots[t][0] == d]
        cur, begin = None, None
        for t in ts + [None]:
            s = assign[t] if t is not None else "__end__"
            if s != cur:
                if cur not in (None, "__end__") and begin is not None:
                    print(f"{d}: {cur:<8} {slots[begin][2]} - {slot_label(slots[t][1]) if t is not None else slot_label(SLOTS_PER_DAY[d])}")
                cur, begin = s, t

    # ---- export to Excel ----
    palette = ["AED6F1", "A9DFBF", "FAD7A0", "F1948A", "D2B4DE", "F9E79F", "A3E4D7", "D5D8DC",
               "85C1E9", "F5B7B1", "ABEBC6", "FDEBD0"]
    color = {s: palette[i % len(palette)] for i, s in enumerate(staff)}
    with pd.ExcelWriter(out_file, engine="openpyxl") as w:
        grid.to_excel(w, sheet_name="Schedule")
        summary.to_excel(w, sheet_name="Hours Summary", index=False)
        ws = w.sheets["Schedule"]
        for row in ws.iter_rows(min_row=2, min_col=2):
            for c in row:
                if c.value in color:
                    c.fill = PatternFill("solid", fgColor=color[c.value])
                    c.font = Font(bold=True)
                c.alignment = Alignment(horizontal="center")
        ws.column_dimensions["A"].width = 10
        for col in "BCDEF":
            ws.column_dimensions[col].width = 12
    print(f"\nSaved: {out_file}")
    return summary, grid


# =====================================================================
# RUN
# =====================================================================
check_inputs([(INPUT_FILE, "returning")])     # stops here if something is wrong
slots, day_slots, staff, ideal, mx, avail, pref = load_data(INPUT_FILE)
print(f"Loaded {len(staff)} staff: {staff}")
print(f"Sum of ideal hrs = {sum(ideal.values())}  |  Sum of max hrs = {sum(mx.values())}  |  Desk hrs = {len(slots)*SLOT_HRS}")

result = solve(slots, day_slots, staff, ideal, mx, avail, pref)
if result:
    assign, hours = result
    summary, grid = report(assign, hours, slots, staff, ideal, mx, pref, OUTPUT_FILE)
