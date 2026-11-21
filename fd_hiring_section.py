# =====================================================================
# HIRING  -  paste this BELOW your existing notebook code
# Compares every possible set of hires and shows the full schedule for the best one.
# Needs the updated fd_scheduler.py code above it (load_data, solve, check_inputs, report, settings).
# =====================================================================
from itertools import combinations

# ---------------- HIRING SETTINGS ----------------
APPLICANT_FILE  = "FD_applicants_template.xlsx"   # second workbook, same template, one sheet per applicant
MAX_HIRES       = 3       # largest hiring set to test (5 applicants, max 3 -> 26 runs)
GUARD_PCT       = 0.75    # flag anyone (returner who lost hours, or new hire) below 75% of their ideal
HIRE_TIME_LIMIT = 60      # seconds per run
HIRE_GAP_TOL    = 0.002   # tighter than before so rows compare fairly


# ---------------- check both workbooks first (stops here if something is wrong) ----------------
cannot_hire = check_inputs([(INPUT_FILE, "returning"), (APPLICANT_FILE, "applicant")])

# ---------------- load returners + applicants ----------------
slots, day_slots, staff, ideal, mx, avail, pref = load_data(INPUT_FILE)
_, _, app_staff, app_ideal, app_mx, app_avail, app_pref = load_data(APPLICANT_FILE)
if set(staff) & set(app_staff):
    raise ValueError(f"Same name in both workbooks: {set(staff) & set(app_staff)}")

all_ideal = {**ideal, **app_ideal}
all_mx    = {**mx, **app_mx}
all_avail = {**avail, **app_avail}
all_pref  = {**pref, **app_pref}

# gaps the returning staff cannot cover, and which applicants can
T_all = range(len(slots))
returner_gaps = [t for t in T_all if sum(avail[s, t] for s in staff) == 0]
print(f"\nSlots no returning staff can cover: {len(returner_gaps)}")
for t in returner_gaps:
    who = [a for a in app_staff if app_avail[a, t]]
    print(f"  {slots[t][0]} {slots[t][2]}  -> applicants available: {who if who else 'NONE'}")


# ---------------- run every hiring set ----------------
hireable = [a for a in app_staff if a not in cannot_hire]
combos = [c for r in range(0, min(MAX_HIRES, len(hireable)) + 1) for c in combinations(hireable, r)]
print(f"\nTesting {len(combos)} hiring sets ({len(hireable)} hireable applicants)...")
raw, skipped = [], []
for i, hired in enumerate(combos, 1):
    staff_set = list(staff) + list(hired)
    min_hrs = {s: MIN_WEEK_HRS for s in staff}
    min_hrs.update({a: min(MIN_HIRE_HRS, all_mx[a]) for a in hired})
    res = solve(slots, day_slots, staff_set, all_ideal, all_mx, all_avail, all_pref,
                time_limit=HIRE_TIME_LIMIT, min_hrs=min_hrs, gap_tol=HIRE_GAP_TOL, quiet=True)
    label = ", ".join(hired) if hired else "(no hires)"
    if res is None:
        skipped.append((label, solve.last_reason))
        print(f"  [{i}/{len(combos)}] {label}: infeasible, skipped"); continue
    print(f"  [{i}/{len(combos)}] {label}: done")
    raw.append({"hired": hired, "staff": staff_set, "assign": res[0], "hours": res[1]})

base = next((r for r in raw if len(r["hired"]) == 0), None)   # the no-hire schedule is the baseline
if base is None:
    print("\nThe no-hire schedule (returning staff only) is infeasible, so nothing can be compared:")
    print(dict(skipped).get("(no hires)", ""))
    raise StopRun

rows = []
for r in raw:
    h, A = r["hours"], r["assign"]
    ret = [s for s in r["staff"] if s in ideal]
    new = [s for s in r["staff"] if s in app_ideal]
    short = lambda s: max(0.0, all_ideal[s] - h[s])
    pct   = lambda s: h[s] / max(all_ideal[s], 1.0)
    covered = [t for t in A if A[t] is not None]
    flags = [f"{s} {pct(s):.0%} of ideal ({h[s] - base['hours'][s]:+.1f}h vs no-hire)"
             for s in ret if pct(s) < GUARD_PCT and h[s] < base["hours"][s] - 0.01]
    flags += [f"{s} (new) {pct(s):.0%} of ideal" for s in new if pct(s) < GUARD_PCT]
    rows.append({
        "Hires": ", ".join(r["hired"]) if r["hired"] else "(no hires)",
        "# Hires": len(r["hired"]),
        "Gap slots": len(A) - len(covered),
        "Worst shortfall %": round(100 * max(short(s) / max(all_ideal[s], 1.0) for s in r["staff"]), 1),
        "Returner shortfall (hrs)": round(sum(short(s) for s in ret), 1),
        "New-hire shortfall (hrs)": round(sum(short(s) for s in new), 1),
        "Total shortfall (hrs)": round(sum(short(s) for s in r["staff"]), 1),
        "Returner hrs vs no-hire": round(sum(h[s] for s in ret) - sum(base["hours"][s] for s in ret), 1),
        "Preferred % of desk hrs": round(100 * sum(all_pref[A[t], t] for t in covered) / max(len(covered), 1), 1),
        "Flags (read me)": "; ".join(flags),
        "_raw": r,
    })

# ranking: fewest gaps -> smallest worst-off shortfall -> smallest total shortfall -> fewest hires
rows.sort(key=lambda d: (d["Gap slots"], round(d["Worst shortfall %"], 0), d["Total shortfall (hrs)"], d["# Hires"]))
SCHEDULES = {}
for rank, d in enumerate(rows, 1):
    d["Rank"] = rank
    SCHEDULES[rank] = d.pop("_raw")

results_df = pd.DataFrame(rows).set_index("Rank")
pd.set_option("display.width", 250); pd.set_option("display.max_columns", 30)
pd.set_option("display.max_colwidth", 80)
print("\n" + "=" * 30 + "\n   HIRING COMPARISON\n" + "=" * 30)
display(results_df)
if skipped:
    print("\nSets that could not be scheduled:")
    for label, reason in skipped:
        print(f"  {label}:\n    " + reason.replace("\n", "\n    "))


# ---------------- full schedule for any row ----------------
def show(rank=1):
    """Full schedule, hours summary and shifts for any row of the table. Also saves it to Excel."""
    r = SCHEDULES[rank]
    print(f"\n########  RANK {rank}: hires = {', '.join(r['hired']) if r['hired'] else '(none)'}  ########")
    return report(r["assign"], r["hours"], slots, r["staff"], all_ideal, all_mx, all_pref,
                  f"fd_schedule_hiring_rank{rank}.xlsx")

summary, grid = show(1)     # best option; call show(2), show(3)... to see the others
