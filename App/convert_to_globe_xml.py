#!/usr/bin/env python3
"""
GloBE Information Return (GIR) XML Generator
Converts Swiss QDMTT Excel calculation template to OECD GIR XML schema format.

Reference: OECD GloBE Information Return (Pillar Two) XML Schema, January 2025
           https://doi.org/10.1787/c594935a-en
"""

import os
import sys
import uuid
import openpyxl
import xml.etree.ElementTree as ET
from datetime import datetime, timezone


# ─── CONFIGURATION ────────────────────────────────────────────────────────────
# Update these values before running for a real filing.
CONFIG = {
    "company_name":   "PLACEHOLDER_COMPANY_AG",
    "tin_value":      "CHE-123456789",
    "tin_issued_by":  "CH",         # ISO 3166-1 Alpha-2
    "tin_type":       "GIR3001",    # GIR3001=TIN, GIR3002=functional equivalent
    "globe_status":   "GIR301",     # GIR301=Constituent Entity
    "rules":          "GIR204",     # GIR204=QDMTT applicable
    "currency":       "CHF",        # ISO 4217
    "jurisdiction":   "CH",         # ISO 3166-1 Alpha-2
    "fas":            "Swiss GAAP FER",
    "cfs_of_upe":     "GIR501",     # GIR501=subparagraph a
    "period_start":   "2024-01-01",
    "period_end":     "2024-12-31",
    "excel_path":     os.path.join(os.path.dirname(os.path.abspath(__file__)), "Calculation File.xlsx"),
    "output_path":    os.path.join(os.path.dirname(os.path.abspath(__file__)), "output", "gir_2024_CH.xml"),
    # Set True to include all GIR adjustment items even when amount is 0
    "include_zero_adjustments": True,
}


# ─── XML NAMESPACES ───────────────────────────────────────────────────────────
# ⚠️  Verify these URIs against the official ESTV-published XSD before filing.
#     ESTV may publish a jurisdiction-specific namespace variant.
GIR_NS = "urn:oecd:ties:gir:v1"
ET.register_namespace("", GIR_NS)

N = f"{{{GIR_NS}}}"  # prefix for namespace-qualified element names


# ─── EXCEL ROW / COLUMN MAPPINGS ─────────────────────────────────────────────
# Primary data column: Column N = jurisdictional totals (=SUM of entity cols F–M)
DATA_COL = "N"

# Row → XML AdjustmentItem enum (NetGlobeIncome adjustments, Article 3.2.x)
INCOME_ADJUSTMENTS: dict[int, str] = {
    238: "GIR2001",  # Net Taxes Expense – Art 3.2.1(a)
    239: "GIR2002",  # Excluded Dividends – Art 3.2.1(b)
    240: "GIR2003",  # Excluded Equity Gain/Loss – Art 3.2.1(c)
    241: "GIR2004",  # Included Revaluation Method Gain/Loss – Art 3.2.1(d)
    242: "GIR2005",  # Gain/loss on disposition excluded under Art 6.3 – 3.2.1(e)
    243: "GIR2006",  # Asymmetric FX Gains/Losses – Art 3.2.1(f)
    244: "GIR2007",  # Policy Disallowed Expenses – Art 3.2.1(g)
    245: "GIR2008",  # Prior Period Errors – Art 3.2.1(h)
    246: "GIR2009",  # Changes in Accounting Principles – Art 3.2.1(h)
    247: "GIR2010",  # Accrued Pension Expense – Art 3.2.1(i)
    248: "GIR2011",  # Debt releases – Art 3.2.1
    249: "GIR2012",  # Stock-based compensation – Art 3.2.2
    250: "GIR2013",  # Arm's length adjustments – Art 3.2.3
    251: "GIR2014",  # Qualified Refundable Tax Credit / MTTC – Art 3.2.4
    252: "GIR2015",  # Election: Gains/losses using realisation principle – Art 3.2.5
    253: "GIR2016",  # Election: Adjusted Asset Gain – Art 3.2.6
    254: "GIR2017",  # Intragroup Financing Arrangement expense – Art 3.2.7
    255: "GIR2018",  # Election: intragroup transactions same jurisdiction – Art 3.2.8
    256: "GIR2019",  # Insurance company taxes charged to policyholders – Art 3.2.9
    257: "GIR2020",  # Additional Tier One Capital – Art 3.2.10
    258: "GIR2021",  # CE joining/leaving MNE Group – Art 3.2.11 & 6.2
    259: "GIR2022",  # Reduction (Flow-through Entity UPE) – Art 3.2.11 & 7.1
    260: "GIR2023",  # Reduction (Deductible Dividend Regime UPE) – Art 3.2.11 & 7.2
    261: "GIR2024",  # Taxable Distribution Method election – Art 3.2.11 & 7.6
    262: "GIR2025",  # International Shipping Income – Art 3.3
    263: "GIR2026",  # Transactions between CEs – Art 9.1.3
}

# Row → XML FinalAdjustedTax enum (AdjustedCoveredTax adjustments, Article 4.x)
# Excel uses GIR24xx internally; XML schema specifies GIR27xx.
COVERED_TAX_ADJUSTMENTS: dict[int, str] = {
    297: "GIR2701",  # Covered Tax accrued as expense – Art 4.1.2(a)        [GIR2401]
    298: "GIR2703",  # Covered Taxes – uncertain tax position – Art 4.1.2(c) [GIR2402]
    299: "GIR2704",  # Qualified RFTC/MTTC reduction – Art 4.1.2(d)          [GIR2403]
    300: "GIR2705",  # Qualified Flow-through Tax Benefits – Art 3.2.1(c)    [GIR2404]
    301: "GIR2706",  # Current tax on excluded income – Art 4.1.3(a)          [GIR2405]
    302: "GIR2707",  # Non-Qualified credits / Other credits – Art 4.1.3(b)   [GIR2406]
    303: "GIR2708",  # Covered Taxes refunded/credited – Art 4.1.3(c)         [GIR2407]
    304: "GIR2709",  # Current tax – uncertain tax position – Art 4.1.3(d)    [GIR2408]
    305: "GIR2710",  # Current tax not paid within 3 years – Art 4.1.3(e)     [GIR2409]
    306: "GIR2711",  # Post-filing adjustments – Art 4.6.1                    [GIR2410]
    307: "GIR2712",  # Covered Taxes – Net Asset Gain/Loss – Art 3.2.6        [GIR2411]
    308: "GIR2713",  # Reduction (Flow-through Entity UPE) – Art 7.1          [GIR2412]
    309: "GIR2714",  # Covered Taxes – Deductible Dividend Regime – Art 7.2.2 [GIR2413]
    310: "GIR2715",  # Deemed Distribution Tax – Art 7.3                      [GIR2414]
    311: "GIR2716",  # Taxable Distribution Method – Art 7.6.2(b)             [GIR2415]
    312: "GIR2717",  # Total Deferred Tax Adjustment Amount – Art 4.4.1(b)    [GIR2416]
    313: "GIR2718",  # Increase/decrease in equity/OCI – Art 4.1.1(c)         [GIR2417]
}

# GIR2702 (GloBE Loss DTA – Art 4.5) has no CE-level row in this template.
# If applicable, add it manually by setting CONFIG["gir2702_amount"].

# GIR2719/2720 come from the OverallComputation section (rows 95–96), not CE-level.
# The OverallComputation section uses column H (first entity column) as input.
ROW_EXCESS_NEG_GENERATED = 95   # GIR2719: Excess Neg Tax Expense generated – Art 4.1.5 & 5.2.1
ROW_EXCESS_NEG_UTILIZED  = 96   # GIR2720: Excess Neg Tax Expense utilised – Art 4.1.5 & 5.2.1
EXCESS_NEG_COL = "H"

# Summary rows (Column N)
ROW_ADJUSTED_FANIL       = 236
ROW_NET_GLOBE_INCOME     = 264
ROW_AGGREGATE_CURR_TAX   = 295
ROW_ADJUSTED_COV_TAX     = 314


# ─── HELPERS ─────────────────────────────────────────────────────────────────

def cell_int(ws, row: int, col: str = DATA_COL) -> int:
    """Return cell value as integer, defaulting to 0."""
    v = ws[f"{col}{row}"].value
    if v is None:
        return 0
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return 0


def fmt_etr(adj_covered_tax: int, net_globe_income: int) -> str:
    """ETR as decimal string (0–1, 4 decimal places) per globe:percentage type."""
    if not net_globe_income:
        return "0.0000"
    rate = adj_covered_tax / net_globe_income
    rate = max(0.0, min(1.0, rate))
    return f"{rate:.4f}"


def sub(parent: ET.Element, tag: str, text=None, **attrib) -> ET.Element:
    """Create a namespace-qualified SubElement with optional text and attributes."""
    el = ET.SubElement(parent, N + tag, attrib)
    if text is not None:
        el.text = str(text)
    return el


# ─── EXCEL READER ────────────────────────────────────────────────────────────

def read_excel(path: str) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["QDMTT 2024"]
    include_zero = CONFIG["include_zero_adjustments"]

    data = {
        "adjusted_fanil":       cell_int(ws, ROW_ADJUSTED_FANIL),
        "net_globe_income":     cell_int(ws, ROW_NET_GLOBE_INCOME),
        "aggregate_curr_tax":   cell_int(ws, ROW_AGGREGATE_CURR_TAX),
        "adjusted_cov_tax":     cell_int(ws, ROW_ADJUSTED_COV_TAX),
        "income_adj":           [],
        "cov_tax_adj":          [],
    }

    for row, gir_code in INCOME_ADJUSTMENTS.items():
        amount = cell_int(ws, row)
        if amount or include_zero:
            data["income_adj"].append((gir_code, amount))

    for row, gir_code in COVERED_TAX_ADJUSTMENTS.items():
        amount = cell_int(ws, row)
        if amount or include_zero:
            data["cov_tax_adj"].append((gir_code, amount))

    # GIR2719 / GIR2720 from OverallComputation section
    gen  = cell_int(ws, ROW_EXCESS_NEG_GENERATED, EXCESS_NEG_COL)
    util = cell_int(ws, ROW_EXCESS_NEG_UTILIZED,  EXCESS_NEG_COL)
    if gen  or include_zero:
        data["cov_tax_adj"].append(("GIR2719", gen))
    if util or include_zero:
        data["cov_tax_adj"].append(("GIR2720", util))

    return data


# ─── XML BUILDER ─────────────────────────────────────────────────────────────

def build_xml(data: dict, cfg: dict) -> ET.Element:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
    msg_ref = (
        f"{cfg['jurisdiction']}2024{cfg['jurisdiction']}"
        f"{str(uuid.uuid4()).replace('-','')[:12].upper()}"
    )

    root = ET.Element(N + "GloBE_Message")

    # ── MessageHeader ─────────────────────────────────────────────────────
    hdr = sub(root, "MessageHeader")
    sub(hdr, "TransmittingCountry", cfg["jurisdiction"])
    sub(hdr, "ReceivingCountry",    cfg["jurisdiction"])
    sub(hdr, "MessageType",         "GIR")
    sub(hdr, "MessageRefID",        msg_ref)
    sub(hdr, "MessageTypeIndic",    "GIR101")   # new information
    sub(hdr, "ReportingPeriod",     cfg["period_end"])
    sub(hdr, "Timestamp",           now)

    # ── GloBE_Body ────────────────────────────────────────────────────────
    body = sub(root, "GloBE_Body")

    # FilingInfo
    fi = sub(body, "FilingInfo")

    filing_ce = sub(fi, "FilingCE")
    id_el = sub(filing_ce, "ID")
    sub(id_el, "Name",           cfg["company_name"])
    sub(id_el, "ResCountryCode", cfg["jurisdiction"])
    sub(id_el, "TIN",            cfg["tin_value"],
               issuedBy=cfg["tin_issued_by"],
               TypeOfTIN=cfg["tin_type"])
    sub(id_el, "Rules",          cfg["rules"])
    sub(id_el, "GlobeStatus",    cfg["globe_status"])

    acct = sub(fi, "AccountingInfo")
    sub(acct, "CFSofUPE", cfg["cfs_of_upe"])
    sub(acct, "FAS",      cfg["fas"])
    sub(acct, "Currency", currCode=cfg["currency"])

    period = sub(fi, "Period")
    sub(period, "Start", cfg["period_start"])
    sub(period, "End",   cfg["period_end"])

    sub(fi, "NameMNE", cfg["company_name"])

    # JurisdictionSection
    jur_sec    = sub(body, "JurisdictionSection")
    globe_tax  = sub(jur_sec, "GloBE_Tax")
    etr        = sub(globe_tax, "ETR")
    etr_status = sub(etr, "ETR_Status")
    etr_comp   = sub(etr_status, "ETR_Computation")
    oc         = sub(etr_comp, "OverallComputation")

    # FANIL / AdjustedFANIL
    sub(oc, "FANIL",        data["adjusted_fanil"])
    sub(oc, "AdjustedFANIL", data["adjusted_fanil"])

    # NetGlobeIncome
    ngi = sub(oc, "NetGlobeIncome")
    sub(ngi, "Total", data["net_globe_income"])
    for gir_code, amount in data["income_adj"]:
        adj = sub(ngi, "Adjustments")
        sub(adj, "Amount",         amount)
        sub(adj, "AdjustmentItem", gir_code)

    # IncomeTaxExpense (aggregate current tax before deferred tax adjustments)
    sub(oc, "IncomeTaxExpense", data["aggregate_curr_tax"])

    # ETRRate
    sub(oc, "ETRRate", fmt_etr(data["adjusted_cov_tax"], data["net_globe_income"]))

    # TopUpTaxPercentage — Validation (required); 0.0000 when no top-up tax applies
    sub(oc, "TopUpTaxPercentage", "0.0000")

    # AdjustedCoveredTax
    act = sub(oc, "AdjustedCoveredTax")
    sub(act, "Total",               data["adjusted_cov_tax"])
    sub(act, "AggregrateCurrentTax", data["aggregate_curr_tax"])
    for gir_code, amount in data["cov_tax_adj"]:
        adj = sub(act, "Adjustments")
        sub(adj, "Amount",         amount)
        sub(adj, "AdjustmentItem", gir_code)

    return root


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main() -> None:
    excel_path  = CONFIG["excel_path"]
    output_path = CONFIG["output_path"]

    if not os.path.exists(excel_path):
        sys.exit(f"ERROR: Excel file not found: {excel_path}")

    print(f"Reading:  {excel_path}")
    data = read_excel(excel_path)

    print(f"  AdjustedFANIL:         {data['adjusted_fanil']:>15,}")
    print(f"  NetGlobeIncome:        {data['net_globe_income']:>15,}")
    print(f"  AggregrateCurrentTax:  {data['aggregate_curr_tax']:>15,}")
    print(f"  AdjustedCoveredTax:    {data['adjusted_cov_tax']:>15,}")
    print(f"  ETRRate:               {fmt_etr(data['adjusted_cov_tax'], data['net_globe_income'])}")
    print(f"  Income adjustments:    {len(data['income_adj'])}")
    print(f"  Tax adjustments:       {len(data['cov_tax_adj'])}")

    root = build_xml(data, CONFIG)
    ET.indent(root, space="  ")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    tree = ET.ElementTree(root)
    tree.write(output_path, xml_declaration=True, encoding="utf-8")
    print(f"\nXML written to: {output_path}")


if __name__ == "__main__":
    main()
