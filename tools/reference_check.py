"""
Independent reference: re-implements the routing rules in pandas and compares
the result with what the VBA macro produced (read from a saved workbook).

    python tools/reference_check.py InvoiceBackup_RESULT.xlsx

Checks, for every tab: same lines, same amount, same intra-EU split; plus that
the CHECK sheet reports OK and the ALL sheet was not modified.
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]


def norm(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return " ".join(str(v).replace(" ", " ").split()).upper()


def clean_sheet(s: str) -> str:
    for ch in '/\\?*[]:\'':
        s = s.replace(ch, " ")
    s = norm(s)
    return s[:31].strip()


def expected(template: Path) -> pd.DataFrame:
    x = pd.ExcelFile(template)
    data = x.parse("ALL")
    ents = x.parse("CONFIG_ENTITIES")
    terms = x.parse("CONFIG_TERMS")
    rules = x.parse("CONFIG_RULES")
    gen = dict(zip(x.parse("CONFIG_GENERAL").Setting, x.parse("CONFIG_GENERAL").Value))
    eu = {norm(c) for c in str(gen["EU Countries"]).split(",") if norm(c)}
    inhouse = norm(gen["In-house Carrier Keyword"])
    suffix_in = norm(gen["In-house Tab Suffix"]) or "GROUP"
    term_map = {norm(c): int(d) for c, d in zip(terms["Carrier Name"], terms["PayDays"])}

    def map_entity(e: str) -> str:
        if not e:
            return "UNASSIGNED"
        for _, r in ents.iterrows():
            m, t = norm(r["Match Text"]), norm(r["Match Type"])
            if (t == "CONTAINS" and m in e) or (t != "CONTAINS" and e == m):
                return norm(r["Tab Name"])
        return e

    def rule_suffix(base, o, d):
        for _, r in rules.iterrows():
            if norm(r["Entity Tab"]) != base:
                continue
            os_ = {norm(c) for c in str(r["Origin Countries"]).split(",") if norm(c)}
            ds = {norm(c) for c in str(r["Destination Countries"]).split(",") if norm(c)}
            if (not os_ or o in os_) and (not ds or d in ds):
                return norm(r["Tab Suffix"])
        return ""

    out = []
    for _, r in data.iterrows():
        base = map_entity(norm(r[gen["Entity Column"]]))
        carrier = norm(r[gen["Carrier Column"]])
        o, d = norm(r[gen["Origin Column"]]), norm(r[gen["Destination Column"]])
        suf = ""
        if base != "UNASSIGNED":
            suf = rule_suffix(base, o, d)
            if not suf:
                if inhouse and inhouse in carrier:
                    suf = suffix_in
                elif carrier in term_map:
                    suf = f"{term_map[carrier]} DAYS"
        tab = clean_sheet(f"{base} - {suf}" if suf else base)
        amt = float(r[gen["Amount Column"]])
        intra = o in eu and d in eu
        out.append((tab, amt, amt if intra else 0.0))
    df = pd.DataFrame(out, columns=["tab", "amount", "intra"])
    return df.groupby("tab").agg(lines=("amount", "size"), amount=("amount", "sum"),
                                 intra=("intra", "sum")).round(2)


def produced(result: Path) -> tuple[pd.DataFrame, dict, list[str]]:
    wb = openpyxl.load_workbook(result, data_only=True)
    s = wb["SUMMARY"]
    rows = [r for r in s.iter_rows(min_row=2, values_only=True) if r[0] and r[0] != "TOTAL"]
    summ = pd.DataFrame([(r[0], r[3], r[4], r[5]) for r in rows],
                        columns=["tab", "lines", "amount", "intra"]).set_index("tab").round(2)
    ck = wb["CHECK"]
    check = {ck.cell(row=r, column=1).value: ck.cell(row=r, column=5).value for r in range(6, 11)}
    check["Overall"] = ck["B3"].value
    # also count lines physically present in each generated tab
    physical = {}
    for tab in summ.index:
        ws = wb[tab]
        physical[tab] = ws.max_row - 1
    summ["physical_lines"] = pd.Series(physical)
    return summ, check, wb.sheetnames


def main(result: Path):
    exp = expected(ROOT / "InvoiceBackup_TEMPLATE.xlsx")
    got, check, sheets = produced(result)
    cmp = exp.join(got, lsuffix="_ref", rsuffix="_vba", how="outer")
    cmp["ok"] = ((cmp.lines_ref == cmp.lines_vba) & (cmp.lines_ref == cmp.physical_lines)
                 & ((cmp.amount_ref - cmp.amount_vba).abs() < 0.01)
                 & ((cmp.intra_ref - cmp.intra_vba).abs() < 0.01))
    print(cmp.to_string())
    print("\nCHECK sheet:", check)
    generated = [s for s in sheets if s in exp.index]
    print(f"\nTabs expected {len(exp)}, in SUMMARY {len(got)}, in workbook {len(generated)}; "
          f"duplicate sheet names: {len(sheets) - len(set(sheets))}")
    all_ok = bool(cmp.ok.all()) and len(generated) == len(exp)
    print("RESULT:", "ALL MATCH" if all_ok else "MISMATCH")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main(Path(sys.argv[1])))
