# Front Desk Scheduler

An optimization model (mixed-integer linear program, solved with Gurobi) that builds a weekly staffing schedule for a university learning center front desk and helps decide which applicants to hire.

**Plain-language walkthrough of the model (variables, objective, constraints, final formulations):**
https://Bundc24.github.io/workforce-scheduling-optimization/

## The problem

The front desk is open Monday to Thursday 9am to 5pm and Friday 9am to 3pm, with exactly one student worker at the desk at a time. That is 76 half-hour blocks (38 hours) to fill.

Each worker submits an Excel sheet with their ideal weekly hours, their maximum weekly hours, and a grid marking each half-hour as **P** (preferred), **A** (available) or **Class** (busy).

The schedule must:

- never place someone in a block they cannot work
- never exceed anyone's maximum hours
- use shifts of 2 to 5 hours, with at most one shift per person per day
- give every returning worker at least one shift

and, as far as possible:

- cover every block
- bring everyone close to their ideal hours, sharing any shortfall fairly
- use blocks people marked as preferred

## Approach

**Part 1: Scheduling.** A MILP with binary variables for "person s works block t", shift starts, and uncovered gaps. The objective combines gaps, worst-off shortfall (a min-max term), total shortfall, overage and preferences, weighted in that order of priority.

**Part 2: Hiring.** For each possible set of applicants (up to 3 hires), the Part 1 model is solved on the combined team. The resulting schedules are compared in a table (gaps, shortfalls, hours lost by returning staff, preferred-hours share) and ranked, and the full schedule for the best set is shown. The equivalent single-MILP version (hire variables with big-M links) is written out on the walkthrough page.

Before solving, the code checks every input sheet and stops with a list of problems (unreadable cells, ideal above max, a person with no legal 2-hour shift, and so on).

## Repository contents

| File | Description |
|---|---|
| `index.html` | Walkthrough page explaining the problem and both formulations |
| `fd_scheduler.py` | Loads the sheets, runs input checks, solves Part 1, prints and exports the schedule |
| `fd_hiring_section.py` | Part 2: compares hiring sets and shows the best schedule |
| `FD_availability_template.xlsx` | Example workbook for returning staff (fictional names) |
| `FD_applicants_template.xlsx` | Example workbook for applicants (fictional names) |

## How to run

1. Install the requirements. Gurobi needs a license (free for academics):

   ```
   pip install pandas openpyxl gurobipy
   ```

2. Put the two Excel workbooks in the same folder as the code, and set `INPUT_FILE` and `APPLICANT_FILE` at the top of the scripts if your file names differ.
3. Run `fd_scheduler.py` to build the returning-staff schedule.
4. Run `fd_hiring_section.py` **in the same Python session** afterwards, because it reuses the functions defined in the first file. The simplest way is Jupyter: paste `fd_scheduler.py` into one cell and run it, then paste `fd_hiring_section.py` into the next cell.

Results are printed and saved as Excel files (`fd_schedule_output.xlsx` and `fd_schedule_hiring_rank1.xlsx`).

## Input format

One sheet per person. Cells `B1:B3` hold name, ideal hours and max hours, and the availability grid starts at row 6 (columns B to F are Mon to Fri, 16 half-hour rows from 9:00a).

| Cell | Meaning |
|---|---|
| `P` | Preferred |
| `A` or blank | Available |
| `Class` (filled red) | Cannot work |

## Limitations

- The objective weights (gap, fairness, shortfall, overage, preference) were chosen by judgment and can be tuned at the top of `fd_scheduler.py`.
- Many schedules can score equally, so a small change to one sheet can reshuffle other people's shifts.
- The hiring comparison tests every set of up to `MAX_HIRES` applicants, so the number of solves grows quickly with a large applicant pool.
- All names in the example workbooks are fictional.

## Built with

Python, Gurobi (`gurobipy`), pandas, openpyxl.
