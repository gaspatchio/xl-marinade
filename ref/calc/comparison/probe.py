"""Function-coverage and rounding probe on a synthetic workbook, both engines.
Expected values are documented Excel results or computed independently in Python."""

import json
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import numpy_financial as nf
import openpyxl
from common import same
from scipy.stats import norm as N

out = Path(sys.argv[1])
out.mkdir(exist_ok=True)
wb = openpyxl.Workbook()
d = wb.active
d.title = "Data"
for i, v in enumerate([10, 20, 30, 40, 50], 1):
    d.cell(i, 1, v)
    d.cell(i, 2, v * 1.5)
    d.cell(i, 3, "abcde"[i - 1])
for i, v in enumerate([-1000, 300, 400, 500, 200], 1):
    d.cell(i, 5, v)
d["G1"] = 0.07
d["G2"] = 2.675
d["G3"] = 1.005
d["G4"] = 0.1
d["G5"] = 0.2
p = wb.create_sheet("Probe")
cf = [-1000, 300, 400, 500, 200]
T = [  # (label, formula, expected)
    ("IF", '=IF(Data!A1>5,"y","n")', "y"),
    ("IFS", "=IFS(Data!A1>50,1,Data!A1>5,2)", 2),
    ("IFERROR", "=IFERROR(1/0,-1)", -1),
    ("AND/OR/NOT", "=AND(TRUE,OR(FALSE,NOT(FALSE)))", True),
    ("CHOOSE", "=CHOOSE(2,10,20,30)", 20),
    ("VLOOKUP exact", "=VLOOKUP(30,Data!A1:C5,3,FALSE)", "c"),
    ("VLOOKUP approx", "=VLOOKUP(35,Data!A1:B5,2,TRUE)", 45),
    ("HLOOKUP", "=HLOOKUP(10,Data!A1:B5,3,FALSE)", 30),
    ("INDEX/MATCH", "=INDEX(Data!B1:B5,MATCH(40,Data!A1:A5,0))", 60),
    ("MATCH approx", "=MATCH(35,Data!A1:A5,1)", 3),
    ("LOOKUP", "=LOOKUP(35,Data!A1:A5,Data!B1:B5)", 45),
    ("XLOOKUP", "=XLOOKUP(20,Data!A1:A5,Data!B1:B5)", 30),
    ("OFFSET", "=SUM(OFFSET(Data!A1,1,0,3,1))", 90),
    ("INDIRECT", '=INDIRECT("Data!A"&3)', 30),
    ("SUM", "=SUM(Data!A1:A5)", 150),
    ("SUMPRODUCT", "=SUMPRODUCT(Data!A1:A5,Data!B1:B5)", 8250),
    ("SUMPRODUCT cond", "=SUMPRODUCT((Data!A1:A5>20)*Data!B1:B5)", 180),
    ("SUMIF", '=SUMIF(Data!A1:A5,">20",Data!B1:B5)', 180),
    ("SUMIFS", '=SUMIFS(Data!B1:B5,Data!A1:A5,">10",Data!A1:A5,"<50")', 135),
    ("COUNT", "=COUNT(Data!A1:C5)", 10),
    ("COUNTA", "=COUNTA(Data!A1:C5)", 15),
    ("COUNTIF", '=COUNTIF(Data!A1:A5,">=30")', 3),
    ("COUNTIFS", '=COUNTIFS(Data!A1:A5,">10",Data!B1:B5,"<70")', 3),
    ("AVERAGE", "=AVERAGE(Data!A1:A5)", 30),
    ("AVERAGEIF", '=AVERAGEIF(Data!A1:A5,">20")', 40),
    ("MIN/MAX", "=MAX(Data!A1:A5)-MIN(Data!A1:A5)", 40),
    ("ROUNDUP", "=ROUNDUP(2.341,2)", 2.35),
    ("ROUNDDOWN", "=ROUNDDOWN(-2.349,2)", -2.34),
    ("INT neg", "=INT(-2.5)", -3),
    ("TRUNC", "=TRUNC(-2.5)", -2),
    ("MOD neg", "=MOD(-7,3)", 2),
    ("CEILING", "=CEILING(2.31,0.05)", 2.35),
    ("FLOOR", "=FLOOR(2.34,0.05)", 2.3),
    ("MROUND", "=MROUND(17,5)", 15),
    ("POWER/EXP/LN", "=LN(EXP(POWER(2,3)))", 8),
    ("SQRT/ABS", "=SQRT(ABS(-16))", 4),
    ("LOG10", "=LOG10(1000)", 3),
    ("NPV", "=NPV(Data!G1,Data!E2:E5)", float(nf.npv(0.07, [0] + cf[1:]))),
    ("IRR", "=IRR(Data!E1:E5)", float(nf.irr(cf))),
    ("PMT", "=PMT(0.05/12,360,200000)", float(nf.pmt(0.05 / 12, 360, 200000))),
    ("PV", "=PV(0.05,10,-100)", float(nf.pv(0.05, 10, -100, 0))),
    ("FV", "=FV(0.05,10,-100)", float(nf.fv(0.05, 10, -100, 0))),
    ("NPER", "=NPER(0.05,-100,1000)", float(nf.nper(0.05, -100, 1000, 0))),
    ("RATE", "=RATE(10,-100,800)", float(nf.rate(10, -100, 800, 0))),
    ("XNPV", "=XNPV(0.07,Data!E1:E2,Data!H1:H2)", -1000 + 300 / 1.07 ** (366 / 365)),
    ("NORM.DIST", "=NORM.DIST(1.2,0.5,2,TRUE)", float(N.cdf(1.2, 0.5, 2))),
    ("NORMDIST pdf", "=NORMDIST(1.2,0.5,2,FALSE)", float(N.pdf(1.2, 0.5, 2))),
    ("NORM.S.DIST", "=NORM.S.DIST(0.3,TRUE)", float(N.cdf(0.3))),
    ("NORMSDIST", "=NORMSDIST(-1.1)", float(N.cdf(-1.1))),
    ("NORM.S.INV", "=NORM.S.INV(0.975)", float(N.ppf(0.975))),
    ("NORMSINV", "=NORMSINV(0.1)", float(N.ppf(0.1))),
    ("NORM.INV", "=NORM.INV(0.9,5,2)", float(N.ppf(0.9, 5, 2))),
    ("STDEV", "=STDEV(Data!A1:A5)", math.sqrt(250)),
    ("VAR.P", "=VAR.P(Data!A1:A5)", 200),
    ("MEDIAN", "=MEDIAN(Data!A1:A5)", 30),
    ("PERCENTILE", "=PERCENTILE(Data!A1:A5,0.9)", 46),
    ("LARGE/SMALL", "=LARGE(Data!A1:A5,2)+SMALL(Data!A1:A5,2)", 60),
    ("RANK", "=RANK(40,Data!A1:A5)", 2),
    ("DATE/YEAR", "=YEAR(DATE(2024,2,29))", 2024),
    ("EDATE", "=EDATE(DATE(2024,1,31),1)-DATE(2024,1,1)", 59),
    ("EOMONTH", "=DAY(EOMONTH(DATE(2023,2,10),0))", 28),
    ("YEARFRAC", "=YEARFRAC(DATE(2020,1,1),DATE(2021,7,1),1)", None),
    ("DATEDIF", '=DATEDIF(DATE(2000,3,15),DATE(2024,3,14),"y")', 23),
    ("DAYS360", "=DAYS360(DATE(2024,1,31),DATE(2024,3,31))", 60),
    ("date serial", "=DATE(2024,1,1)", 45292),
    ("TEXT", '=TEXT(1234.5,"0.00")', "1234.50"),
    ("CONCAT &", '=LEFT("abcdef",2)&MID("abcdef",3,2)&RIGHT("abcdef",1)', "abcdf"),
    ("LEN/TRIM/UPPER", '=LEN(TRIM(UPPER(" ab ")))', 2),
    ("VALUE", '=VALUE("12.5")*2', 25),
    ("ISNUMBER/ISBLANK", "=AND(ISNUMBER(Data!A1),ISBLANK(Data!Z9))", True),
    ("NA/ISNA", "=ISNA(MATCH(99,Data!A1:A5,0))", True),
    ("text+num coercion", '="3"+4', 7),
    ("bool arithmetic", "=TRUE+TRUE", 2),
    ("blank as zero", "=Data!Z9+5", 5),
    ("div0 error", "=1/0", "#DIV/0!"),
    ("pct operator", "=50%*200", 100),
    ("precedence -2^2", "=-2^2", 4),
    ("string compare ci", '="ABC"="abc"', True),
    ("SUM ignores text", "=SUM(Data!A1:C5)", 375),
    ("SUBTOTAL", "=SUBTOTAL(9,Data!A1:A5)", 150),
    ("AGGREGATE", "=AGGREGATE(9,6,Data!A1:A5)", 150),
    ("MMULT via SUMPRODUCT", "=SUMPRODUCT(MMULT(Data!A1:B2,Data!A1:B2))", 3000),
    ("TRANSPOSE in SUMPRODUCT", "=SUMPRODUCT(TRANSPOSE(Data!A1:A5))", 150),
    ("RAND range", "=AND(RAND()>=0,RAND()<1)", True),
    ("TODAY is number", "=ISNUMBER(TODAY())", True),
    ("unknown UDF", "=MyUdf(3)", "#NAME?"),
    # rounding: literal ties whose Excel result is documented (half away from zero on the decimal value)
    ("ROUND 2.675,2", "=ROUND(2.675,2)", 2.68),
    ("ROUND 1.005,2", "=ROUND(1.005,2)", 1.01),
    ("ROUND 0.285,2", "=ROUND(0.285,2)", 0.29),
    ("ROUND 1.45,1", "=ROUND(1.45,1)", 1.5),
    ("ROUND 8.345,2", "=ROUND(8.345,2)", 8.35),
    ("ROUND -2.675,2", "=ROUND(-2.675,2)", -2.68),
    ("ROUND 2.5,0", "=ROUND(2.5,0)", 3),
    ("ROUND -0.5,0", "=ROUND(-0.5,0)", -1),
    ("ROUND 1234.5,-1", "=ROUND(1235,-1)", 1240),
    ("ROUND cell 2.675", "=ROUND(Data!G2,2)", 2.68),
    ("ROUND cell 1.005", "=ROUND(Data!G3,2)", 1.01),
    # computed ties: Excel's answer is NOT asserted here (None) -- engines are only compared with each other
    ("ROUND (0.1+0.2)*5,1 [computed]", "=ROUND((Data!G4+Data!G5)*0.5,2)", None),
    ("ROUND 1.15*3,2 [computed]", "=ROUND(1.15*3,2)", None),
    ("ROUND 0.615*100,0 [computed]", "=ROUND(0.615*100,0)", None),
    ("ROUND 4.35*100,0 [computed]", "=ROUND(4.35*100,0)", None),
    ("ROUND 1.1*1.1*100,0... [computed]", "=ROUND(1.255*100,0)", None),
    ("0.1+0.2=0.3", "=(0.1+0.2)=0.3", None),
    ("1-0.9-0.1", "=1-0.9-0.1", None),
]
import datetime

d["H1"] = datetime.date(2024, 1, 1)
d["H2"] = datetime.date(2025, 1, 1)
for i, (lab, f, exp) in enumerate(T, 1):
    p.cell(i, 1, lab)
    p.cell(i, 2, f)
path = out / "probe.xlsx"
wb.save(path)


def ironcalc_vals():
    import ironcalc

    m = ironcalc.load_from_xlsx(str(path), "en", "UTC", "en")
    m.evaluate()
    return [m.get_cell_value(1, i, 2) for i in range(1, len(T) + 1)]


def formulas_vals():
    import formulas

    xl = formulas.ExcelModel().loads(str(path)).finish()
    sol = xl.calculate()
    vals = []
    for i in range(1, len(T) + 1):
        k = f"'[probe.xlsx]PROBE'!B{i}"
        try:
            v = sol[k].value[0, 0]
        except Exception:
            v = "#MISSING"
        vals.append(
            v.item()
            if hasattr(v, "item")
            else (str(v) if not isinstance(v, (int, float, bool, str, type(None))) else v)
        )
    return vals


res = {}
for name, fn in (("ironcalc", ironcalc_vals), ("formulas", formulas_vals)):
    try:
        res[name] = fn()
    except Exception as e:
        res[name] = [f"#LOAD:{type(e).__name__}:{str(e)[:80]}"] * len(T)
rows = []
for i, (lab, f, exp) in enumerate(T):
    r = {"label": lab, "formula": f, "expected": exp}
    for e in res:
        g = res[e][i]
        r[e] = g
        r[e + "_ok"] = None if exp is None else bool(same(exp, g))
    r["engines_agree"] = bool(same(res["ironcalc"][i], res["formulas"][i]))
    rows.append(r)
json.dump(rows, open(out / "probe.json", "w", encoding="utf-8"), indent=1, default=str)
for e in res:
    bad = [r for r in rows if r[e + "_ok"] is False]
    print(
        f"\n== {e}: {sum(1 for r in rows if r[e + '_ok'])} of {sum(1 for r in rows if r['expected'] is not None)} checked cases correct; wrong/unsupported:"
    )
    for r in bad:
        print(f"   {r['label']:28} expected {r['expected']!r:22} got {r[e]!r}")
print("\n== unasserted cases (engine vs engine):")
for r in rows:
    if r["expected"] is None:
        print(
            f"   {r['label']:36} ironcalc={r['ironcalc']!r:24} formulas={r['formulas']!r:24} agree={r['engines_agree']}"
        )
