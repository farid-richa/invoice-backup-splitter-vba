"""
Build the sample workbook InvoiceBackup_TEMPLATE.xlsx with fictional data.

Everything is invented: the client group "Valoria", its legal entities, the
carriers, shipments and amounts. The data is deliberately a bit messy
(trailing spaces, non-breaking spaces, two spellings of one entity, a few
rows with no entity) because real exports look like that and the macro has
to cope.

    python tools/build_workbook.py
"""
from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUT = Path(__file__).resolve().parents[1] / "InvoiceBackup_TEMPLATE.xlsx"
SEED = 7
N_LINES = 420

NAVY = "0C182B"
HDR_FONT = Font(bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor=NAVY)
CFG_FILL = PatternFill("solid", fgColor="1F9E85")

# raw spellings as they come out of the billing system -> weight
ENTITIES = [
    ("VALORIA CONSUMER CARE AG", 0.30),
    ("Valoria Consumer Care AG ", 0.05),               # trailing space
    ("VALORIA HEALTH AG", 0.18),
    ("VALORIA HEALTH AKTIENGESELLSCHAFT", 0.07),       # same entity, long form
    ("VALORIA OY", 0.15),
    ("VALORIA STEINFELD GMBH", 0.10),
    ("VALORIA IBERICA SL", 0.12),                 # non-breaking space
    ("", 0.03),                                        # missing entity
]
ENTITY_COUNTRY = {
    "VALORIA CONSUMER CARE AG": "CHE", "VALORIA HEALTH AG": "DEU", "VALORIA OY": "FIN",
    "VALORIA STEINFELD GMBH": "DEU", "VALORIA IBERICA SL": "ESP",
}

CARRIERS = [  # id, name, mode
    ("LEVEXP", "LEVANTE EXPRESS SL", "ROAD-LTL"),
    ("MESHAU", "MESETA HAULAGE SA", "ROAD-FTL"),
    ("NORCAR", "NORTE CARGO GMBH", "ROAD-FTL"),
    ("COSFRT", "COSTA FREIGHT BV", "ROAD-LTL"),
    ("RIOTRP", "RIO TRANSPORTES LDA", "ROAD-FTL"),
    ("BALRAIL", "BALTIC RAIL LOGISTICS AB", "RAIL"),
    ("NSTFRT", "NORTHSTAR FREIGHT GMBH", "ROAD-FTL"),   # in-house group carrier
    ("NSTPAR", "NORTHSTAR PARCEL AB", "PARCEL"),       # in-house group carrier
]
CARRIER_WEIGHTS = [0.17, 0.14, 0.14, 0.12, 0.10, 0.08, 0.15, 0.10]

COUNTRIES = ["DEU", "FRA", "ESP", "ITA", "NLD", "BEL", "POL", "AUT", "SWE", "NOR", "FIN",
             "CHE", "GBR", "DNK", "PRT"]

REJECTION_REASONS = ["Rate above contract", "Duplicate invoice", "Missing proof of delivery",
                     "Unagreed accessorial", "Wrong bill-to entity"]

HEADERS = ["Shipment ID", "Invoice Number", "Invoice Date", "Invoice Status", "Rejection Reason",
           "Carrier ID", "Carrier Name", "Bill-To Entity", "Origin Country", "Destination Country",
           "Mode", "Chargeable Weight KG", "Base Cost EUR", "Fuel EUR", "Accessorials EUR",
           "Amount EUR"]

EU = ["AUT", "BEL", "BGR", "HRV", "CYP", "CZE", "DNK", "EST", "FIN", "FRA", "DEU", "GRC", "HUN",
      "IRL", "ITA", "LVA", "LTU", "LUX", "MLT", "NLD", "POL", "PRT", "ROU", "SVK", "SVN", "ESP",
      "SWE"]


def pick(rng, items, weights):
    return rng.choices(items, weights=weights, k=1)[0]


def build_rows(rng: random.Random) -> list[list]:
    rows = []
    start = date(2026, 9, 1)
    for i in range(N_LINES):
        entity = pick(rng, [e for e, _ in ENTITIES], [w for _, w in ENTITIES])
        cid, cname, mode = pick(rng, CARRIERS, CARRIER_WEIGHTS)
        clean = " ".join(entity.replace(" ", " ").split()).upper()
        home = ENTITY_COUNTRY.get(clean.replace("AKTIENGESELLSCHAFT", "AG"), rng.choice(COUNTRIES))

        # a handful of Nordic cross-border lanes for Valoria Oy (special rule)
        if clean == "VALORIA OY" and rng.random() < 0.35:
            origin, dest = rng.choice(["SWE", "NOR"]), "FIN"
        elif rng.random() < 0.5:
            origin, dest = home, rng.choice([c for c in COUNTRIES if c != home])
        else:
            origin, dest = rng.choice([c for c in COUNTRIES if c != home]), home

        weight = round(rng.lognormvariate(7.2, 1.0), 1)
        if mode == "PARCEL":
            weight = round(rng.uniform(2, 60), 1)
        base = round(max(45.0, weight * rng.uniform(0.08, 0.22)), 2)
        if mode == "RAIL":
            base = round(base * 3.5, 2)
        fuel = round(base * rng.uniform(0.10, 0.16), 2)
        acc = round(rng.choice([0, 0, 0, 0, 35, 45, 60, 90]) * 1.0, 2)
        total = round(base + fuel + acc, 2)

        status = pick(rng, ["APPROVED", "PENDING", "REJECTED"], [0.68, 0.22, 0.10])
        reason = rng.choice(REJECTION_REASONS) if status == "REJECTED" else ""
        inv_date = start + timedelta(days=rng.randint(0, 45))

        rows.append([
            f"SHP{2026000 + i:07d}", f"{cid}-{rng.randint(10000, 99999)}", inv_date, status, reason,
            cid, cname, entity, origin, dest, mode, weight, base, fuel, acc, total,
        ])
    # shipments are sometimes split over two invoice lines with the same carrier name
    # written slightly differently (extra space) - Norm() must treat them as one carrier
    for r in rng.sample(rows, 6):
        r[6] = r[6] + " "
    return rows


def style_header(ws, ncols, fill=HDR_FILL):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HDR_FONT
        cell.fill = fill
        cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"


def autosize(ws, min_w=10, max_w=45):
    for col in ws.columns:
        length = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(min_w, min(max_w, length + 2))


def main():
    rng = random.Random(SEED)
    wb = Workbook()

    # ---------- README ----------
    ws = wb.active
    ws.title = "README"
    ws.sheet_view.showGridLines = False
    lines = [
        ("Invoice Backup Splitter", Font(size=16, bold=True, color=NAVY)),
        ("All data in this workbook is fictional.", Font(italic=True, color="5B6B7F")),
        ("", None),
        ("1. Paste the billing export into the ALL sheet (keep the header row).", None),
        ("2. Review the CONFIG_* sheets: entities, payment terms, special routing rules, general settings.", None),
        ("3. Run GenerateEntityTabs  ->  one tab per entity / rule, plus SUMMARY and CHECK.", None),
        ("4. Run ExportPendingInvoices or ExportRejectedInvoices  ->  one file per carrier in the exports folder.", None),
        ("", None),
        ("Routing priority for each line:", Font(bold=True)),
        ("   a) special rule in CONFIG_RULES (entity + origin + destination)", None),
        ("   b) in-house group carrier (keyword in CONFIG_GENERAL)", None),
        ("   c) carrier payment terms in CONFIG_TERMS", None),
        ("   d) the entity's base tab", None),
    ]
    for i, (text, font) in enumerate(lines, start=2):
        c = ws.cell(row=i, column=2, value=text)
        if font:
            c.font = font
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 110

    # ---------- ALL ----------
    ws = wb.create_sheet("ALL")
    ws.append(HEADERS)
    for r in build_rows(rng):
        ws.append(r)
    style_header(ws, len(HEADERS))
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        row[0].number_format = "dd/mm/yyyy"
    for col in range(12, 17):
        for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            row[0].number_format = "#,##0.00"
    autosize(ws)
    ws.auto_filter.ref = ws.dimensions

    # ---------- CONFIG_ENTITIES ----------
    ws = wb.create_sheet("CONFIG_ENTITIES")
    ws.append(["Match Text", "Match Type", "Tab Name", "Display Name"])
    for row in [
        ["VALORIA CONSUMER CARE", "CONTAINS", "VALORIA CONSUMER CARE", "Valoria Consumer Care AG"],
        ["VALORIA HEALTH AG", "EXACT", "VALORIA HEALTH", "Valoria Health AG"],
        ["VALORIA HEALTH AKTIENGESELLSCHAFT", "EXACT", "VALORIA HEALTH", "Valoria Health AG"],
        ["VALORIA OY", "EXACT", "VALORIA OY", "Valoria Oy"],
        ["STEINFELD", "CONTAINS", "VALORIA STEINFELD", "Valoria Steinfeld GmbH"],
        ["VALORIA IBERICA", "CONTAINS", "VALORIA IBERICA", "Valoria Iberica SL"],
    ]:
        ws.append(row)
    style_header(ws, 4, CFG_FILL)
    autosize(ws)

    # ---------- CONFIG_TERMS ----------
    ws = wb.create_sheet("CONFIG_TERMS")
    ws.append(["Carrier Name", "PayDays"])
    for row in [["MESETA HAULAGE SA", 60], ["BALTIC RAIL LOGISTICS AB", 75], ["RIO TRANSPORTES LDA", 45]]:
        ws.append(row)
    style_header(ws, 2, CFG_FILL)
    autosize(ws)

    # ---------- CONFIG_RULES ----------
    ws = wb.create_sheet("CONFIG_RULES")
    ws.append(["Rule Name", "Entity Tab", "Origin Countries", "Destination Countries", "Tab Suffix"])
    ws.append(["Nordic cross-border", "VALORIA OY", "SWE,NOR", "FIN", "CROSS-BORDER"])
    style_header(ws, 5, CFG_FILL)
    autosize(ws)

    # ---------- CONFIG_GENERAL ----------
    ws = wb.create_sheet("CONFIG_GENERAL")
    ws.append(["Setting", "Value"])
    for row in [
        ["Entity Column", "Bill-To Entity"],
        ["Carrier Column", "Carrier Name"],
        ["Carrier ID Column", "Carrier ID"],
        ["Origin Column", "Origin Country"],
        ["Destination Column", "Destination Country"],
        ["Status Column", "Invoice Status"],
        ["Amount Column", "Amount EUR"],
        ["In-house Carrier Keyword", "NORTHSTAR"],
        ["In-house Tab Suffix", "GROUP"],
        ["EU Countries", ",".join(EU)],
        ["Export Folder", "exports"],
    ]:
        ws.append(row)
    style_header(ws, 2, CFG_FILL)
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 60

    wb.save(OUT)
    print(f"wrote {OUT} ({N_LINES} lines)")


if __name__ == "__main__":
    main()
