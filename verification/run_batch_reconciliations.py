"""Run every workbook's in-workbook VBA batch reconciliation and save state.

For each workbook: open the .xlsm, run the reconciliation macro over the
FULL population (BatchLimit blank), read the Batch_Results summary, save
the workbook (so the committed artifact carries the full-run evidence),
and write the canonical report JSON. Runs under the machine-wide Excel
lock. Usage: run_batch_reconciliations.py [BasicTerm_S SmithWilson ...]
"""

import json
import subprocess
import sys
from datetime import date
from pathlib import Path

import xlwings as xw

from excel_lock import hold


def osa(script, seconds=120):
    res = subprocess.run(["osascript", "-e", script],
                         capture_output=True, text=True, timeout=seconds + 60)
    if res.returncode != 0:
        raise RuntimeError(f"osascript failed: {res.stderr.strip()}")
    return res.stdout.strip()


def run_macro_with_timeout(wb_name, macro, seconds=3600):
    """Run a macro on Excel's DEFAULT instance with a long Apple-event timeout.

    Two hard-won rules (marinade-notes P5/P10): xlwings' wb.macro() times out
    at 60s and cannot be extended, and long batch macros run inside an
    xlwings-spawned `newinstance` can wedge in the CalculationState wait.
    The workbook must be opened in, and the macro run on, the default
    instance via osascript.
    """
    osa(f'''
    with timeout of {seconds} seconds
        tell application "Microsoft Excel"
            run VB macro "{wb_name}!{macro}"
        end tell
    end timeout''', seconds)

REPO = Path(__file__).resolve().parents[1]
REPORTS = REPO / "verification" / "reports"

JOBS = {
    "BasicTerm_S": {
        "wb": "workbooks/BasicTerm_S/BasicTerm_S.xlsm",
        "macro": "RunAllModelPoints",
        "report": "basicterm_s_batch_reconciliation.json",
        "reference": "verification/reference/basicterm_all_points.csv "
                     "(lifelib basiclife.BasicTerm_M, all 10,000 points)",
        "quantities": "office premium and PVs of premiums, claims, expenses, "
                      "commissions and net cash flow (6 per point)",
        "summary": {"points_run": ("C4", int), "reconciled": ("C5", int),
                    "mismatches": ("C6", int), "max_abs_diff_premium": ("C7", float),
                    "max_rel_diff_pv_components": ("C8", float),
                    "runtime_seconds": ("C9", float), "verdict": ("C10", str)},
        "err_cell": "D10",
    },
    "BasicTerm_SE": {
        "wb": "workbooks/BasicTerm_SE/BasicTerm_SE.xlsm",
        "macro": "RunAllModelPoints",
        "report": "basicterm_se_batch_reconciliation.json",
        "reference": "verification/reference/basicterm_se_all_points.csv "
                     "(lifelib basiclife.BasicTerm_ME, all 10,000 points)",
        "quantities": "office premium and PVs of premiums, claims, expenses, "
                      "commissions and net cash flow (6 per point)",
        "summary": {"points_run": ("C4", int), "reconciled": ("C5", int),
                    "mismatches": ("C6", int), "max_abs_diff_premium": ("C7", float),
                    "max_rel_diff_pv_components": ("C8", float),
                    "runtime_seconds": ("C9", float), "verdict": ("C10", str)},
        "err_cell": "D10",
    },
    "CashValue_SE": {
        "wb": "workbooks/CashValue_SE/CashValue_SE.xlsm",
        "macro": "RunAllModelPoints",
        "report": "cashvalue_se_batch_reconciliation.json",
        "reference": "verification/reference/cashvalue_se_all_runs.csv "
                     "(lifelib savings.CashValue_SE, every point x scenario)",
        "quantities": "PVs of premiums, claims, expenses, commissions, investment "
                      "income, AV change and net cash flow (7 per run)",
        "summary": {"runs": ("C4", int), "reconciled": ("C5", int),
                    "mismatches": ("C6", int),
                    "max_rel_diff_pv_premiums": ("C7", float),
                    "max_rel_diff_pv_components": ("C8", float),
                    "runtime_seconds": ("C9", float), "verdict": ("C10", str)},
        "err_cell": "D10",
    },
    "SmithWilson": {
        "wb": "workbooks/SmithWilson/SmithWilson.xlsm",
        "macro": "RunCurveReconciliation",
        "report": "smithwilson_batch_reconciliation.json",
        "reference": "verification/reference/smithwilson.json "
                     "(lifelib smithwilson: zeta vector, P and R curves)",
        "quantities": "zeta (25), P (65) and R (65) - 155 values",
        "summary": {"values_compared": ("C4", int), "reconciled": ("C5", int),
                    "mismatches": ("C6", int), "max_abs_diff_zeta": ("C7", float),
                    "max_diff_P_over_max1": ("C8", float),
                    "max_diff_R_over_max1": ("C9", float),
                    "runtime_seconds": ("C10", float), "verdict": ("C11", str)},
        "err_cell": "D11",
    },
    "Solvency2": {
        "wb": "workbooks/Solvency2/Solvency2.xlsm",
        "macro": "RunAllPolicies",
        "report": "solvency2_batch_reconciliation.json",
        "reference": "verification/reference/solvency2_all_policies.csv "
                     "(lifelib solvency2, every policy, t0=0, scen 1)",
        "quantities": "SCR, Life(mort/longev/lapse/exps), the three lapse "
                      "sub-shock charges and base NAV (9 per policy)",
        "summary": {"policies_run": ("C4", int), "reconciled": ("C5", int),
                    "mismatches": ("C6", int), "max_rel_diff_scr": ("C7", float),
                    "max_rel_diff_life_charges": ("C8", float),
                    "max_rel_diff_nav": ("C9", float),
                    "runtime_seconds": ("C10", float), "verdict": ("C11", str)},
        "err_cell": "D11",
    },
}


def run_job(name, job):
    path = REPO / job["wb"]
    # freshly built xlsm files carry no cached values, so Excel fully
    # recalculates on open - the open reply can take minutes
    osa(f'''
    with timeout of 1800 seconds
        tell application "Microsoft Excel"
            open POSIX file "{path}"
        end tell
    end timeout''', 1800)
    wb = xw.Book(str(path))  # attach to the default instance
    try:
        wb.names["BatchLimit"].refers_to_range.value = None
    except Exception:
        pass  # SmithWilson has no BatchLimit
    run_macro_with_timeout(path.name, job["macro"])
    br = wb.sheets["Batch_Results"]
    out = {
        "workbook": job["wb"],
        "method": f"in-workbook VBA macro {job['macro']} (BatchRunner.bas), "
                  "full run saved on Batch_Results sheet",
        "reference": job["reference"],
        "quantities_reconciled": job["quantities"],
        "date": date.today().isoformat(),
        "tolerance_relative": 1e-06,
    }
    for key, (cell, cast) in job["summary"].items():
        v = br.range(cell).value
        out[key] = cast(v) if v is not None else None
    err = br.range(job["err_cell"]).value
    if err:
        out["error"] = err
    wb.save()
    wb.close()
    (REPORTS / job["report"]).write_text(json.dumps(out, indent=1))
    print(f"{name}: {out['verdict']} -> {job['report']} "
          f"({out.get('runtime_seconds')}s)")
    return out["verdict"] == "PASS" and not err


def main():
    names = sys.argv[1:] or list(JOBS)
    ok = True
    with hold("batch reconciliations (all workbooks)"):
        try:
            for name in names:
                ok = run_job(name, JOBS[name]) and ok
        finally:
            subprocess.run(["osascript", "-e",
                            'tell application "Microsoft Excel" to quit'],
                           capture_output=True)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
