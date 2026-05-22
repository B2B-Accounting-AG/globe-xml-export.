#!/usr/bin/env python3
"""
GloBE XML Export App
Streamlit web UI for converting the Swiss QDMTT Excel template to OECD GIR XML.
Run: streamlit run globe_xml_app.py
"""

import io
import logging
import os
import re
import uuid
import zipfile
import openpyxl
import xml.etree.ElementTree as ET
from datetime import date, datetime, timezone

from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives import padding as sym_padding
from cryptography.hazmat.primitives.asymmetric import padding as asym_padding
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives.serialization import load_pem_public_key
from cryptography.x509 import load_pem_x509_certificate

import base64
from PIL import Image
import streamlit as st

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)

VERSION = "1.5.3"

# ─── XML SETUP ───────────────────────────────────────────────────────────────

GIR_NS = "urn:oecd:ties:globe:v2"        # confirmed from GLOBEXML_v1.0.xsd targetNamespace
STF_NS = "urn:oecd:ties:globestf:v5"    # DocSpec children (DocTypeIndic, DocRefId)
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
N = "{" + GIR_NS + "}"
S = "{" + STF_NS + "}"
ET.register_namespace("globe", GIR_NS)
ET.register_namespace("stf", STF_NS)
ET.register_namespace("xsi", XSI_NS)


# ─── MAPPINGS ────────────────────────────────────────────────────────────────

DATA_COL = "N"

INCOME_ADJUSTMENTS: dict[int, str] = {
    238: "GIR2001", 239: "GIR2002", 240: "GIR2003", 241: "GIR2004",
    242: "GIR2005", 243: "GIR2006", 244: "GIR2007", 245: "GIR2008",
    246: "GIR2009", 247: "GIR2010", 248: "GIR2011", 249: "GIR2012",
    250: "GIR2013", 251: "GIR2014", 252: "GIR2015", 253: "GIR2016",
    254: "GIR2017", 255: "GIR2018", 256: "GIR2019", 257: "GIR2020",
    258: "GIR2021", 259: "GIR2022", 260: "GIR2023", 261: "GIR2024",
    262: "GIR2025", 263: "GIR2026",
}

COVERED_TAX_ADJUSTMENTS: dict[int, str] = {
    297: "GIR2701", 298: "GIR2703", 299: "GIR2704", 300: "GIR2705",
    301: "GIR2706", 302: "GIR2707", 303: "GIR2708", 304: "GIR2709",
    305: "GIR2710", 306: "GIR2711", 307: "GIR2712", 308: "GIR2713",
    309: "GIR2714", 310: "GIR2715", 311: "GIR2716", 312: "GIR2717",
    313: "GIR2718",
}

GIR_INCOME_LABELS: dict[str, str] = {
    "GIR2001": "Net Taxes Expense (Art. 3.2.1a)",
    "GIR2002": "Excluded Dividends (Art. 3.2.1b)",
    "GIR2003": "Excluded Equity Gain or Loss (Art. 3.2.1c)",
    "GIR2004": "Included Revaluation Method Gain or Loss (Art. 3.2.1d)",
    "GIR2005": "Gain/Loss on Disposal of Assets excluded under Art. 6.3 (Art. 3.2.1e)",
    "GIR2006": "Asymmetric Foreign Currency Gains or Losses (Art. 3.2.1f)",
    "GIR2007": "Policy Disallowed Expenses (Art. 3.2.1g)",
    "GIR2008": "Prior Period Errors (Art. 3.2.1h)",
    "GIR2009": "Changes in Accounting Principles (Art. 3.2.1h)",
    "GIR2010": "Accrued Pension Expense (Art. 3.2.1i)",
    "GIR2011": "Debt Releases (Art. 3.2.1)",
    "GIR2012": "Stock-based Compensation (Art. 3.2.2)",
    "GIR2013": "Arm's Length Adjustments (Art. 3.2.3)",
    "GIR2014": "QRTC / Marketable Transferable Tax Credit (Art. 3.2.4)",
    "GIR2015": "Election for Gains/Losses – Realisation Principle (Art. 3.2.5)",
    "GIR2016": "Election for Adjusted Asset Gain (Art. 3.2.6)",
    "GIR2017": "Intragroup Financing Arrangement Expense (Art. 3.2.7)",
    "GIR2018": "Election for Intragroup Transactions – Same Jurisdiction (Art. 3.2.8)",
    "GIR2019": "Insurance Company Taxes Charged to Policyholders (Art. 3.2.9)",
    "GIR2020": "AT1/RT1 Capital Distribution Adjustments (Art. 3.2.10)",
    "GIR2021": "CE Joining/Leaving MNE Group (Art. 3.2.11 & 6.2)",
    "GIR2022": "Reduction of GloBE Income – UPE Flow-through Entity (Art. 3.2.11 & 7.1)",
    "GIR2023": "Reduction of GloBE Income – UPE Deductible Dividend Regime (Art. 3.2.11 & 7.2)",
    "GIR2024": "Taxable Distribution Method Election (Art. 3.2.11 & 7.6)",
    "GIR2025": "International Shipping Income (Art. 3.3)",
    "GIR2026": "Transactions between Constituent Entities (Art. 9.1.3)",
}

GIR_TAX_LABELS: dict[str, str] = {
    "GIR2701": "Covered Tax Accrued as Expense (Art. 4.1.2a)",
    "GIR2702": "GloBE Loss Deferred Tax Asset (Art. 4.1.2b & 4.5.3)",
    "GIR2703": "Covered Taxes – Uncertain Tax Position Prior Year (Art. 4.1.2c)",
    "GIR2704": "QRTC / Marketable Transferable Tax Credit – Current Tax Reduction (Art. 4.1.2d)",
    "GIR2705": "Qualified Flow-through Tax Benefits (Art. 3.2.1c)",
    "GIR2706": "Current Tax on Excluded Income (Art. 4.1.3a)",
    "GIR2707": "Non-QRTC / Other Tax Credits (Art. 4.1.3b)",
    "GIR2708": "Covered Taxes Refunded or Credited (Art. 4.1.3c)",
    "GIR2709": "Current Tax – Uncertain Tax Position (Art. 4.1.3d)",
    "GIR2710": "Current Tax Not Expected Paid within 3 Years (Art. 4.1.3e)",
    "GIR2711": "Post-filing Adjustments (Art. 4.6.1)",
    "GIR2712": "Covered Taxes – Net Asset Gain/Loss (Art. 3.2.6)",
    "GIR2713": "Reduction – UPE Flow-through Entity (Art. 7.1)",
    "GIR2714": "Covered Taxes – UPE Deductible Dividend Regime (Art. 7.2.2)",
    "GIR2715": "Deemed Distribution Tax (Art. 7.3)",
    "GIR2716": "Taxable Distribution Method Election (Art. 7.6b)",
    "GIR2717": "Total Deferred Tax Adjustment Amount (Art. 4.4.1b)",
    "GIR2718": "Covered Taxes in Equity / OCI (Art. 4.1.1c)",
    "GIR2719": "Excess Negative Tax Expense Carry Forward Generated (Art. 4.1.5 & 5.2.1)",
    "GIR2720": "Excess Negative Tax Expense Carry Forward Utilized (Art. 4.1.5 & 5.2.1)",
}

UPE_RULES_OPTIONS = ["GIR201", "GIR202", "GIR203", "GIR204", "GIR205"]
UPE_RULES_LABELS: dict[str, str] = {
    "GIR201": "GIR201 — QIIR (other jurisdictions only)",
    "GIR202": "GIR202 — QIIR (other + parent jurisdiction)",
    "GIR203": "GIR203 — QUTPR",
    "GIR204": "GIR204 — QDMTT",
    "GIR205": "GIR205 — Not applicable",
}

UPE_GLOBE_STATUS_OPTIONS = [
    "GIR301", "GIR302", "GIR303", "GIR304", "GIR305", "GIR306",
    "GIR307", "GIR308", "GIR309", "GIR310", "GIR311", "GIR312",
    "GIR313", "GIR314", "GIR315", "GIR316", "GIR317", "GIR318",
]
UPE_GLOBE_STATUS_LABELS: dict[str, str] = {
    "GIR301": "GIR301 — Constituent Entity",
    "GIR302": "GIR302 — Flow-Through Entity – Tax Transparent",
    "GIR303": "GIR303 — Flow-Through Entity – Reverse Hybrid",
    "GIR304": "GIR304 — Hybrid Entity",
    "GIR305": "GIR305 — Permanent Establishment",
    "GIR306": "GIR306 — Main Entity",
    "GIR307": "GIR307 — Minority-Owned Parent Entity",
    "GIR308": "GIR308 — Minority-Owned Subsidiary",
    "GIR309": "GIR309 — Minority-Owned Constituent Entity",
    "GIR310": "GIR310 — Investment Entity",
    "GIR311": "GIR311 — Insurance Investment Entity",
    "GIR312": "GIR312 — Securitisation Entity",
    "GIR313": "GIR313 — JV",
    "GIR314": "GIR314 — JV Subsidiary",
    "GIR315": "GIR315 — Non-Material Constituent Entity",
    "GIR316": "GIR316 — Excluded Entity",
    "GIR317": "GIR317 — Parent Entity (Art. 10.3.5)",
    "GIR318": "GIR318 — Non-group Member",
}

ROW_EXCESS_NEG_GENERATED = 95
ROW_EXCESS_NEG_UTILIZED  = 96
EXCESS_NEG_COL = "H"
ROW_ADJUSTED_FANIL     = 236
ROW_NET_GLOBE_INCOME   = 264
ROW_AGGREGATE_CURR_TAX = 295
ROW_ADJUSTED_COV_TAX   = 314


# ─── UI OPTIONS ──────────────────────────────────────────────────────────────

CURRENCIES = [
    "CHF", "EUR", "USD", "GBP", "JPY", "AUD", "CAD", "SEK", "NOK", "DKK",
    "SGD", "HKD", "NZD", "CNY", "INR", "BRL", "MXN", "KRW", "ZAR", "RUB",
    "PLN", "CZK", "HUF", "RON", "BGN", "ISK", "TRY", "SAR", "AED",
    "ILS", "THB", "IDR", "MYR", "PHP", "CLP", "ARS", "COP", "PEN",
    "EGP", "NGN", "KES", "MAD",
]

FAS_OPTIONS = [
    "Swiss GAAP FER",
    "IFRS",
    "US GAAP",
    "UK GAAP",
    "HGB",
    "Local GAAP",
]

_COUNTRIES = [
    ("AD", "Andorra"), ("AE", "United Arab Emirates"), ("AL", "Albania"),
    ("AM", "Armenia"), ("AO", "Angola"), ("AR", "Argentina"), ("AT", "Austria"),
    ("AU", "Australia"), ("AZ", "Azerbaijan"), ("BA", "Bosnia and Herzegovina"),
    ("BB", "Barbados"), ("BE", "Belgium"), ("BG", "Bulgaria"), ("BH", "Bahrain"),
    ("BM", "Bermuda"), ("BR", "Brazil"), ("BS", "Bahamas"), ("BW", "Botswana"),
    ("BY", "Belarus"), ("BZ", "Belize"), ("CA", "Canada"), ("CH", "Switzerland"),
    ("CL", "Chile"), ("CN", "China"), ("CO", "Colombia"), ("CR", "Costa Rica"),
    ("CY", "Cyprus"), ("CZ", "Czech Republic"), ("DE", "Germany"), ("DK", "Denmark"),
    ("DZ", "Algeria"), ("EE", "Estonia"), ("EG", "Egypt"), ("ES", "Spain"),
    ("FI", "Finland"), ("FR", "France"), ("GB", "United Kingdom"), ("GE", "Georgia"),
    ("GH", "Ghana"), ("GI", "Gibraltar"), ("GR", "Greece"), ("GT", "Guatemala"),
    ("HK", "Hong Kong"), ("HR", "Croatia"), ("HU", "Hungary"), ("ID", "Indonesia"),
    ("IE", "Ireland"), ("IL", "Israel"), ("IN", "India"), ("IS", "Iceland"),
    ("IT", "Italy"), ("JE", "Jersey"), ("JM", "Jamaica"), ("JO", "Jordan"),
    ("JP", "Japan"), ("KE", "Kenya"), ("KG", "Kyrgyzstan"), ("KR", "South Korea"),
    ("KW", "Kuwait"), ("KZ", "Kazakhstan"), ("LB", "Lebanon"), ("LI", "Liechtenstein"),
    ("LT", "Lithuania"), ("LU", "Luxembourg"), ("LV", "Latvia"), ("MA", "Morocco"),
    ("MC", "Monaco"), ("MD", "Moldova"), ("ME", "Montenegro"), ("MK", "North Macedonia"),
    ("MT", "Malta"), ("MU", "Mauritius"), ("MX", "Mexico"), ("MY", "Malaysia"),
    ("NA", "Namibia"), ("NG", "Nigeria"), ("NL", "Netherlands"), ("NO", "Norway"),
    ("NZ", "New Zealand"), ("OM", "Oman"), ("PA", "Panama"), ("PE", "Peru"),
    ("PH", "Philippines"), ("PK", "Pakistan"), ("PL", "Poland"), ("PT", "Portugal"),
    ("QA", "Qatar"), ("RO", "Romania"), ("RS", "Serbia"), ("RU", "Russia"),
    ("SA", "Saudi Arabia"), ("SE", "Sweden"), ("SG", "Singapore"), ("SI", "Slovenia"),
    ("SK", "Slovakia"), ("SM", "San Marino"), ("TH", "Thailand"), ("TN", "Tunisia"),
    ("TR", "Turkey"), ("TT", "Trinidad and Tobago"), ("TW", "Taiwan"),
    ("UA", "Ukraine"), ("UG", "Uganda"), ("US", "United States"), ("UY", "Uruguay"),
    ("UZ", "Uzbekistan"), ("VN", "Vietnam"), ("ZA", "South Africa"),
    ("ZM", "Zambia"), ("ZW", "Zimbabwe"),
]
COUNTRY_DISPLAY = [f"{code} – {name}" for code, name in _COUNTRIES]


def _country_idx(code: str) -> int:
    for i, (c, _) in enumerate(_COUNTRIES):
        if c == code:
            return i
    return 0


# ─── HELPERS ─────────────────────────────────────────────────────────────────

def cell_int(ws, row: int, col: str = DATA_COL) -> int:
    v = ws[f"{col}{row}"].value
    if v is None:
        return 0
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return 0


def fmt_etr(adj_covered_tax: int, net_globe_income: int) -> str:
    if not net_globe_income:
        return "0.0000"
    rate = max(0.0, min(1.0, adj_covered_tax / net_globe_income))
    return f"{rate:.4f}"


def sub(parent: ET.Element, tag: str, text=None, **attrib) -> ET.Element:
    el = ET.SubElement(parent, N + tag, attrib)
    if text is not None:
        el.text = str(text)
    return el


# ─── CORE CONVERSION ─────────────────────────────────────────────────────────

def read_excel(file_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb["QDMTT 2024"]

    data = {
        "adjusted_fanil":     cell_int(ws, ROW_ADJUSTED_FANIL),
        "net_globe_income":   cell_int(ws, ROW_NET_GLOBE_INCOME),
        "aggregate_curr_tax": cell_int(ws, ROW_AGGREGATE_CURR_TAX),
        "adjusted_cov_tax":   cell_int(ws, ROW_ADJUSTED_COV_TAX),
        "income_adj":         [],
        "cov_tax_adj":        [],
    }

    for row, gir_code in INCOME_ADJUSTMENTS.items():
        data["income_adj"].append((gir_code, cell_int(ws, row)))

    for row, gir_code in COVERED_TAX_ADJUSTMENTS.items():
        data["cov_tax_adj"].append((gir_code, cell_int(ws, row)))

    gen  = cell_int(ws, ROW_EXCESS_NEG_GENERATED, EXCESS_NEG_COL)
    util = cell_int(ws, ROW_EXCESS_NEG_UTILIZED,  EXCESS_NEG_COL)
    data["cov_tax_adj"].append(("GIR2719", gen))
    data["cov_tax_adj"].append(("GIR2720", util))

    return data


def build_xml(data: dict, cfg: dict, test_mode: bool = False) -> str:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
    year = cfg["period_end"][:4]
    msg_ref = f"{cfg['jurisdiction']}{year}{cfg['jurisdiction']}{str(uuid.uuid4())}"

    root = ET.Element(N + "GLOBE_OECD", {
        "version": "1.0",
        "{" + XSI_NS + "}schemaLocation": "urn:oecd:ties:globe:v2 GLOBEXML_v1.0.xsd",
    })

    hdr = sub(root, "MessageSpec")
    sub(hdr, "SendingEntityIN",     cfg["tin_value"])
    sub(hdr, "TransmittingCountry", cfg["jurisdiction"])
    sub(hdr, "ReceivingCountry",    cfg["jurisdiction"])
    sub(hdr, "MessageType",         "GIR")
    sub(hdr, "MessageRefId",        msg_ref)
    sub(hdr, "MessageTypeIndic",    "GIR101")
    sub(hdr, "ReportingPeriod",     cfg["period_end"])
    sub(hdr, "Timestamp",           now)

    body = sub(root, "GLOBEBody")
    fi   = sub(body, "FilingInfo")

    filing_ce = sub(fi, "FilingCE")
    sub(filing_ce, "ResCountryCode", cfg["jurisdiction"])
    sub(filing_ce, "Name",           cfg["company_name"])
    sub(filing_ce, "TIN",            cfg["tin_value"],
                   issuedBy=cfg["tin_issued_by"], TypeOfTIN=cfg["tin_type"])
    sub(filing_ce, "Role",           cfg["reporting_role"])

    acct = sub(fi, "AccountingInfo")
    sub(acct, "CFSofUPE", cfg["cfs_of_upe"])
    sub(acct, "FAS",      cfg["fas"])
    sub(acct, "Currency", cfg["currency"])

    period = sub(fi, "Period")
    sub(period, "Start", cfg["period_start"])
    sub(period, "End",   cfg["period_end"])
    sub(fi, "NameMNE", cfg["company_name"])

    doc_type_indic = "OECD10" if test_mode else "OECD1"
    # TEMP v1.5.3: ESTV CTS validator misreads OECD10 as OECD0 in non-FilingInfo sections (60013/60014).
    # Use OECD11 in GeneralSection/JurisdictionSection to bypass. Revert to doc_type_indic once ESTV fixes CTS.
    doc_type_indic_sections = "OECD11" if test_mode else "OECD1"

    fi_doc = sub(fi, "DocSpec")
    ET.SubElement(fi_doc, S + "DocTypeIndic").text = doc_type_indic
    ET.SubElement(fi_doc, S + "DocRefId").text = f"{cfg['jurisdiction']}{year}-{str(uuid.uuid4())}"

    gen_sec = sub(body, "GeneralSection")
    sub(gen_sec, "RecJurCode", cfg["rec_jur_code"])
    corp = sub(gen_sec, "CorporateStructure")
    upe_el = sub(corp, "UPE")
    other_upe = sub(upe_el, "OtherUPE")
    id_el = sub(other_upe, "ID")
    sub(id_el, "Name", cfg["company_name"])
    sub(id_el, "ResCountryCode", cfg["jurisdiction"])
    upe_tin = ET.SubElement(id_el, N + "TIN",
                            {"issuedBy": cfg["tin_issued_by"], "TypeOfTIN": cfg["tin_type"]})
    upe_tin.text = cfg["tin_value"]
    sub(id_el, "Rules", cfg["upe_rules"])
    sub(id_el, "GlobeStatus", cfg["upe_globe_status"])
    gen_doc = sub(gen_sec, "DocSpec")
    ET.SubElement(gen_doc, S + "DocTypeIndic").text = doc_type_indic_sections
    ET.SubElement(gen_doc, S + "DocRefId").text = f"{cfg['jurisdiction']}{year}-{str(uuid.uuid4())}"

    jur_sec = sub(body, "JurisdictionSection")
    sub(jur_sec, "RecJurCode",  cfg["rec_jur_code"])
    sub(jur_sec, "Jurisdiction", cfg["rec_jur_code"])

    globe_tax  = sub(jur_sec, "GLoBETax")
    etr        = sub(globe_tax, "ETR")
    etr_status = sub(etr, "ETRStatus")
    etr_comp   = sub(etr_status, "ETRComputation")
    oc         = sub(etr_comp, "OverallComputation")

    sub(oc, "FANIL",         data["adjusted_fanil"])
    sub(oc, "AdjustedFANIL", data["adjusted_fanil"])

    ngi = sub(oc, "NetGlobeIncome")
    sub(ngi, "Total", data["net_globe_income"])
    for gir_code, amount in data["income_adj"]:
        if amount != 0:
            adj = sub(ngi, "Adjustments")
            sub(adj, "Amount",         amount)
            sub(adj, "AdjustmentItem", gir_code)

    sub(oc, "IncomeTaxExpense",    data["aggregate_curr_tax"])
    sub(oc, "ETRRate",             fmt_etr(data["adjusted_cov_tax"], data["net_globe_income"]))
    sub(oc, "TopUpTaxPercentage",  "0.0000")

    act = sub(oc, "AdjustedCoveredTax")
    sub(act, "Total",                data["adjusted_cov_tax"])
    sub(act, "AggregrateCurrentTax", data["aggregate_curr_tax"])
    for gir_code, amount in data["cov_tax_adj"]:
        if amount != 0:
            adj = sub(act, "Adjustments")
            sub(adj, "Amount",         amount)
            sub(adj, "AdjustmentItem", gir_code)

    # ExcessProfits = NetGlobeIncome - SubstanceExclusion (SBIE=0 for CH QDMTT)
    sub(oc, "ExcessProfits", data["net_globe_income"])

    qdmtt = sub(oc, "QDMTT")
    sub(qdmtt, "FAS",            cfg["fas"])
    sub(qdmtt, "Amount",         data["adjusted_cov_tax"])
    sub(qdmtt, "SBIEAvailable",  "false")
    sub(qdmtt, "DeMinAvailable", "false")
    sub(qdmtt, "Currency",       cfg["currency"])

    sub(oc, "TopUpTax", 0)

    ente = sub(oc, "ExcessNegTaxExpense")
    sub(ente, "PriorYearBalance", 0)
    sub(ente, "GeneratedInRFY",   0)
    sub(ente, "UtilizedInRFY",    0)
    sub(ente, "Remaining",        0)

    jur_doc = sub(jur_sec, "DocSpec")
    ET.SubElement(jur_doc, S + "DocTypeIndic").text = doc_type_indic_sections
    ET.SubElement(jur_doc, S + "DocRefId").text = f"{cfg['jurisdiction']}{year}-{str(uuid.uuid4())}"

    ET.indent(root, space="  ")
    raw = ET.tostring(root, encoding="unicode")
    return f"<?xml version='1.0' encoding='utf-8'?>\n{raw}"


# ─── ENCRYPTION ──────────────────────────────────────────────────────────────

def encrypt_for_estv(xml_str: str, pem_bytes: bytes) -> bytes:
    """
    Packages xml_str into the ESTV-required encrypted zip:
      Payload.xml → Payload.zip (compressed) → AES-256-CBC encrypted → "Payload"
      AES key + IV (48 bytes) → RSA PKCS#1 v1.5 encrypted → "Key"
      Final zip contains "Payload" + "Key"
    """
    # Step 1: compress XML into Payload.zip
    inner_buf = io.BytesIO()
    with zipfile.ZipFile(inner_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("Payload.xml", xml_str.encode("utf-8"))
    payload_zip = inner_buf.getvalue()

    # Step 2: AES-256 CBC encrypt the zip
    aes_key = os.urandom(32)
    iv      = os.urandom(16)
    padder  = sym_padding.PKCS7(128).padder()
    padded  = padder.update(payload_zip) + padder.finalize()
    cipher  = Cipher(algorithms.AES(aes_key), modes.CBC(iv), backend=default_backend())
    enc     = cipher.encryptor()
    encrypted_payload = enc.update(padded) + enc.finalize()

    # Step 3: RSA-encrypt the 48-byte key material (key || iv)
    try:
        pub_key = load_pem_public_key(pem_bytes, backend=default_backend())
    except Exception:
        cert    = load_pem_x509_certificate(pem_bytes, default_backend())
        pub_key = cert.public_key()
    encrypted_key = pub_key.encrypt(aes_key + iv, asym_padding.PKCS1v15())

    # Step 4: bundle into final zip
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_STORED) as zf:
        zf.writestr("Payload", encrypted_payload)
        zf.writestr("Key",     encrypted_key)
    return out_buf.getvalue()


# ─── VALIDATION ──────────────────────────────────────────────────────────────

def validate_xml(xml_str: str) -> list[tuple[str, bool, str]]:
    """Returns list of (label, passed, detail) tuples."""
    results = []

    def check(label, passed, detail=""):
        results.append((label, passed, detail))

    # 1. Well-formed XML
    try:
        root = ET.fromstring(xml_str)
        check("Well-formed XML", True)
    except ET.ParseError as e:
        check("Well-formed XML", False, str(e))
        return results

    def _nsp(path: str) -> str:
        return "/".join(N + s for s in path.split("/"))

    def text(path):
        el = root.find(_nsp(path))
        return el.text.strip() if el is not None and el.text else None

    def findall(path):
        return root.findall(_nsp(path))

    # 2. Root element
    check("Root element (GLOBE_OECD)",
          root.tag == N + "GLOBE_OECD" and root.find(N + "MessageSpec") is not None)

    # 3. MessageSpec — all required fields (incl. Swiss SendingEntityIN)
    hdr_fields = ["TransmittingCountry", "ReceivingCountry", "MessageType",
                  "MessageRefId", "MessageTypeIndic", "ReportingPeriod",
                  "Timestamp", "SendingEntityIN"]
    missing_hdr = [f for f in hdr_fields if text(f"MessageSpec/{f}") is None]
    check("MessageSpec — all required fields (incl. SendingEntityIN)", not missing_hdr,
          f"Missing: {', '.join(missing_hdr)}" if missing_hdr else "")

    # 4. MessageRefId format: CH[0-9]{4}CH...
    msg_ref = text("MessageSpec/MessageRefId")
    check("MessageRefId format (CH[year]CH[uuid])",
          bool(msg_ref and re.match(r"^[A-Z]{2}\d{4}[A-Z]{2}.+", msg_ref)),
          msg_ref or "missing")

    # 5. Timestamp format
    ts = text("MessageSpec/Timestamp")
    check("Timestamp format (YYYY-MM-DDTHH:MM:SS)",
          bool(ts and re.match(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}", ts)),
          ts or "missing")

    # 6. Period dates
    start = text("GLOBEBody/FilingInfo/Period/Start")
    end   = text("GLOBEBody/FilingInfo/Period/End")
    date_ok = bool(
        start and re.match(r"\d{4}-\d{2}-\d{2}$", start) and
        end   and re.match(r"\d{4}-\d{2}-\d{2}$", end)
    )
    check("Period dates (YYYY-MM-DD)", date_ok,
          f"Start: {start}  End: {end}" if not date_ok else "")

    # 7. Company name — not placeholder
    name = text("GLOBEBody/FilingInfo/FilingCE/Name")
    name_ok = bool(name and name != "PLACEHOLDER_COMPANY_AG")
    check("Company name (not placeholder)", name_ok,
          "Still set to PLACEHOLDER_COMPANY_AG" if not name_ok else "")

    # 8. Role in FilingCE (GIR401–GIR405)
    role = text("GLOBEBody/FilingInfo/FilingCE/Role")
    check("FilingCE Role (GIR401–GIR405)",
          bool(role and re.match(r"^GIR40[1-5]$", role)),
          role or "missing")

    # 9. TIN — not placeholder, has required attributes
    tin_el = root.find(_nsp("GLOBEBody/FilingInfo/FilingCE/TIN"))
    tin_val = tin_el.text.strip() if tin_el is not None and tin_el.text else None
    tin_ok = bool(tin_val and tin_val != "CHE-123456789")
    check("TIN (not placeholder)", tin_ok,
          "Still set to CHE-123456789" if not tin_ok else "")
    if tin_el is not None:
        check("TIN attributes (issuedBy + TypeOfTIN)",
              bool(tin_el.get("issuedBy") and tin_el.get("TypeOfTIN")))

    # 10. DocSpec in FilingInfo
    fi_doc = root.find(_nsp("GLOBEBody/FilingInfo/DocSpec"))
    fi_doc_ok = (
        fi_doc is not None and
        fi_doc.find(S + "DocTypeIndic") is not None and
        fi_doc.find(S + "DocRefId") is not None
    )
    check("FilingInfo DocSpec (DocTypeIndic + DocRefId)", fi_doc_ok)

    # 11. RecJurCode in JurisdictionSection
    rec_jur = text("GLOBEBody/JurisdictionSection/RecJurCode")
    check("JurisdictionSection RecJurCode present",
          bool(rec_jur and re.match(r"^[A-Z]{2}$", rec_jur)),
          rec_jur or "missing")

    # 12. DocSpec in JurisdictionSection
    jur_doc = root.find(_nsp("GLOBEBody/JurisdictionSection/DocSpec"))
    jur_doc_ok = (
        jur_doc is not None and
        jur_doc.find(S + "DocTypeIndic") is not None and
        jur_doc.find(S + "DocRefId") is not None
    )
    check("JurisdictionSection DocSpec (DocTypeIndic + DocRefId)", jur_doc_ok)

    # Currency — ISO 4217 text content
    ccy_el = root.find(_nsp("GLOBEBody/FilingInfo/AccountingInfo/Currency"))
    check("Currency (ISO 4217 text)",
          bool(ccy_el is not None and ccy_el.text and len(ccy_el.text.strip()) == 3))

    # OverallComputation — required elements
    oc = ("GLOBEBody/JurisdictionSection/GLoBETax/ETR"
          "/ETRStatus/ETRComputation/OverallComputation")
    oc_fields = ["FANIL", "AdjustedFANIL", "IncomeTaxExpense",
                 "ETRRate", "TopUpTaxPercentage"]
    missing_oc = [f for f in oc_fields if text(f"{oc}/{f}") is None]
    check("OverallComputation — required elements", not missing_oc,
          f"Missing: {', '.join(missing_oc)}" if missing_oc else "")

    # ETRRate — decimal 0–1, 4 decimal places
    etr_val = text(f"{oc}/ETRRate")
    try:
        etr_f  = float(etr_val) if etr_val else None
        etr_ok = (etr_f is not None and 0 <= etr_f <= 1
                  and bool(re.match(r"^\d\.\d{4}$", etr_val)))
    except (ValueError, TypeError):
        etr_ok = False
    check("ETRRate format (0.0000 – 1.0000)", etr_ok, etr_val or "missing")

    # TopUpTaxPercentage format
    tup = text(f"{oc}/TopUpTaxPercentage")
    check("TopUpTaxPercentage format (0.0000)",
          bool(tup and re.match(r"^\d\.\d{4}$", tup)), tup or "missing")

    # NetGlobeIncome — only valid GIR20xx codes (zero-value items are omitted)
    valid_ngi = {f"GIR{2000+i}" for i in range(1, 27)}
    ngi_codes = {el.text for el in findall(
        f"{oc}/NetGlobeIncome/Adjustments/AdjustmentItem") if el.text}
    invalid_ngi = ngi_codes - valid_ngi
    check("NetGlobeIncome — adjustment codes valid (GIR2001–GIR2026)", not invalid_ngi,
          f"Invalid: {', '.join(sorted(invalid_ngi))}" if invalid_ngi else "")

    # AdjustedCoveredTax — only valid GIR27xx codes
    valid_act = {f"GIR27{i:02d}" for i in range(1, 21)}
    act_codes = {el.text for el in findall(
        f"{oc}/AdjustedCoveredTax/Adjustments/AdjustmentItem") if el.text}
    invalid_act = act_codes - valid_act
    check("AdjustedCoveredTax — adjustment codes valid (GIR2701–GIR2720)", not invalid_act,
          f"Invalid: {', '.join(sorted(invalid_act))}" if invalid_act else "")

    # GeneralSection present with CorporateStructure
    gen_sec = root.find(_nsp("GLOBEBody/GeneralSection"))
    gen_ok = (
        gen_sec is not None and
        gen_sec.find(_nsp("CorporateStructure")) is not None and
        gen_sec.find(_nsp("CorporateStructure/UPE")) is not None
    )
    check("GeneralSection present (CorporateStructure / UPE)", gen_ok)

    # All amounts are integers
    non_int = [el.text for el in root.findall(".//" + N + "Amount")
               if el.text and "." in el.text]
    check("All amounts are integers (no decimals)", not non_int,
          f"Non-integer: {non_int[:3]}" if non_int else "")

    return results


# ─── TRANSLATIONS ────────────────────────────────────────────────────────────

T: dict[str, dict[str, str]] = {
    "step1":               {"EN": "1. Upload Excel file",              "DE": "1. Excel-Datei hochladen"},
    "upload_label":        {"EN": "Calculation File (.xlsx or .xlsm)", "DE": "Berechnungsdatei (.xlsx oder .xlsm)"},
    "upload_help":         {"EN": 'Swiss QDMTT calculation template with sheet "QDMTT 2024"',
                            "DE": 'Schweizer QDMTT-Berechnungsvorlage mit Tabellenblatt "QDMTT 2024"'},
    "step2":               {"EN": "2. Company details",                "DE": "2. Unternehmensangaben"},
    "company_name":        {"EN": "Company name",                      "DE": "Firmenname"},
    "tin":                 {"EN": "TIN",                               "DE": "UID/TIN"},
    "tin_issued_by":       {"EN": "TIN issued by (ISO 3166-1 Alpha-2)","DE": "TIN ausgestellt von (ISO 3166-1 Alpha-2)"},
    "jurisdiction":        {"EN": "Jurisdiction (ISO 3166-1 Alpha-2)", "DE": "Jurisdiktion (ISO 3166-1 Alpha-2)"},
    "currency":            {"EN": "Currency (ISO 4217)",               "DE": "Währung (ISO 4217)"},
    "fas_label":           {"EN": "Financial Accounting Standard",     "DE": "Rechnungslegungsstandard"},
    "period_start":        {"EN": "Period start",                      "DE": "Periodenstart"},
    "period_end":          {"EN": "Period end",                        "DE": "Periodenende"},
    "partner_country":     {"EN": "Partner country (RecJurCode)",      "DE": "Partnerstaat (RecJurCode)"},
    "partner_help":        {"EN": "ISO 3166-1 Alpha-2 country code of the partner jurisdiction (must not be CH)",
                            "DE": "ISO 3166-1 Alpha-2 Ländercode des Partnerstaats (darf nicht CH sein)"},
    "advanced":            {"EN": "Advanced options",                  "DE": "Erweiterte Optionen"},
    "filing_role":         {"EN": "Filing role",                       "DE": "Einreichungsrolle"},
    "filing_role_help":    {"EN": "Role as registered in the ESTV ePortal",
                            "DE": "Rolle gemäss Registrierung im ESTV ePortal"},
    "tin_type_label":      {"EN": "TIN type",                         "DE": "TIN-Typ"},
    "tin_type_help":       {"EN": "Type of identifier used as TIN",   "DE": "Art des als TIN verwendeten Identifikators"},
    "cfs_upe":             {"EN": "CFS of UPE",                       "DE": "CFS der UPE"},
    "cfs_upe_help":        {"EN": "Type of Consolidated Financial Statement of the Ultimate Parent Entity",
                            "DE": "Art des konsolidierten Abschlusses der obersten Muttergesellschaft"},
    "submission_mode":     {"EN": "Submission mode",                   "DE": "Einreichungsmodus"},
    "mode_production":     {"EN": "Production (OECD1)",                "DE": "Produktion (OECD1)"},
    "mode_test":           {"EN": "Test / CTS (OECD10)",               "DE": "Test / CTS (OECD10)"},
    "mode_help":           {"EN": "Use Test/CTS for the acceptance portal (eportal-a.admin.ch). Use Production for the live portal (eportal.admin.ch).",
                            "DE": "Test/CTS für das Abnahmeportal (eportal-a.admin.ch), Produktion für das Live-Portal (eportal.admin.ch)."},
    "gir401":              {"EN": "GIR401 — Ultimate Parent Entity (UPE)",    "DE": "GIR401 — Oberste Muttergesellschaft (UPE)"},
    "gir402":              {"EN": "GIR402 — Designated Filing Entity (DFE)",  "DE": "GIR402 — Benannte Einreichungsstelle (DFE)"},
    "gir404":              {"EN": "GIR404 — Constituent Entity (CE)",         "DE": "GIR404 — Untereinheit (CE)"},
    "gir3001":             {"EN": "GIR3001 — Tax Identification Number (TIN)","DE": "GIR3001 — Steueridentifikationsnummer (TIN)"},
    "gir3002":             {"EN": "GIR3002 — Functional equivalent",          "DE": "GIR3002 — Funktionales Äquivalent"},
    "gir501":              {"EN": "GIR501 — Consolidated Financial Statement (subparagraph a)",
                            "DE": "GIR501 — Konsolidierter Abschluss (Buchstabe a)"},
    "gir502":              {"EN": "GIR502 — Combined Financial Statement (subparagraph b)",
                            "DE": "GIR502 — Kombinierter Abschluss (Buchstabe b)"},
    "gir503":              {"EN": "GIR503 — Other",                           "DE": "GIR503 — Sonstiges"},
    "step3":               {"EN": "3. Export",                         "DE": "3. Export"},
    "generate_btn":        {"EN": "Generate XML",                      "DE": "XML generieren"},
    "upload_first":        {"EN": "Please upload an Excel file first.","DE": "Bitte laden Sie zuerst eine Excel-Datei hoch."},
    "spinner_gen":         {"EN": "Reading Excel and building XML…",   "DE": "Excel wird gelesen und XML wird erstellt…"},
    "upload_to_enable":    {"EN": "Upload the Excel file above to enable export.",
                            "DE": "Laden Sie die Excel-Datei oben hoch, um den Export zu aktivieren."},
    "validation_title":    {"EN": "Structural validation",             "DE": "Strukturelle Validierung"},
    "checks_passed":       {"EN": "checks passed",                     "DE": "Prüfungen bestanden"},
    "fix_issues":          {"EN": ("Fix the issues above, then re-generate. "
                                   "Once all checks pass, validate against the official ESTV XSD before submission."),
                            "DE": ("Beheben Sie die oben genannten Probleme und generieren Sie erneut. "
                                   "Sobald alle Prüfungen bestanden sind, validieren Sie gegen das offizielle ESTV XSD vor der Einreichung.")},
    "all_passed":          {"EN": "All structural checks passed.",     "DE": "Alle strukturellen Prüfungen bestanden."},
    "download_xml":        {"EN": "⬇️  Download XML",                 "DE": "⬇️  XML herunterladen"},
    "preview_xml":         {"EN": "Preview XML",                       "DE": "XML-Vorschau"},
    "sheet_not_found":     {"EN": 'Sheet not found: {}. Make sure the file contains a sheet named "QDMTT 2024".',
                            "DE": 'Tabellenblatt nicht gefunden: {}. Stellen Sie sicher, dass die Datei ein Blatt namens "QDMTT 2024" enthält.'},
    "error_msg":           {"EN": "Error: {}",                         "DE": "Fehler: {}"},
    "step4":               {"EN": "4. Encrypt for ESTV",              "DE": "4. Verschlüsselung für ESTV"},
    "step4_caption":       {"EN": ("Upload the ESTV public key (ESTV-PublicKey.pem) from the myESTV portal. "
                                   "The app will produce an encrypted .zip ready to upload directly to the GIR-Applikation."),
                            "DE": ("Laden Sie den öffentlichen ESTV-Schlüssel (ESTV-PublicKey.pem) aus dem myESTV-Portal hoch. "
                                   "Die App erstellt eine verschlüsselte .zip-Datei, die direkt in die GIR-Applikation hochgeladen werden kann.")},
    "pem_label":           {"EN": "ESTV Public Key (.pem)",            "DE": "Öffentlicher ESTV-Schlüssel (.pem)"},
    "encrypt_btn":         {"EN": "Encrypt & Download",               "DE": "Verschlüsseln & Herunterladen"},
    "generate_first":      {"EN": "Generate the XML first (Step 3).", "DE": "Generieren Sie zuerst das XML (Schritt 3)."},
    "upload_pem":          {"EN": "Upload the ESTV public key (.pem) above.",
                            "DE": "Laden Sie oben den öffentlichen ESTV-Schlüssel (.pem) hoch."},
    "spinner_enc":         {"EN": "Encrypting…",                      "DE": "Verschlüsselung läuft…"},
    "download_zip":        {"EN": "⬇️  Download encrypted ZIP",       "DE": "⬇️  Verschlüsselte ZIP herunterladen"},
    "ready_estv":          {"EN": "Ready to upload to myESTV → GIR-Applikation.",
                            "DE": "Bereit zum Hochladen in myESTV → GIR-Applikation."},
    "enc_failed":          {"EN": "Encryption failed: {}",            "DE": "Verschlüsselung fehlgeschlagen: {}"},
    "gen_first_long":      {"EN": "Generate the XML in Step 3 first, then encrypt here.",
                            "DE": "Generieren Sie zuerst das XML in Schritt 3, dann verschlüsseln Sie hier."},
    "bundled_key_info":    {"EN": "Using bundled ESTV public key (encryptor.estv.admin.ch, valid until 2027-02-04).",
                            "DE": "Verwendung des integrierten ESTV-Schlüssels (encryptor.estv.admin.ch, gültig bis 04.02.2027)."},
    "override_key":        {"EN": "Use a different key",              "DE": "Anderen Schlüssel verwenden"},
    "override_active":     {"EN": "Custom key active",                "DE": "Eigener Schlüssel aktiv"},
    "err_jurisdiction":    {"EN": "Jurisdiction must be exactly 2 uppercase letters (e.g. CH)",
                            "DE": "Die Jurisdiktion muss genau 2 Grossbuchstaben sein (z.B. CH)"},
    "err_currency":        {"EN": "Currency must be exactly 3 uppercase letters (e.g. CHF)",
                            "DE": "Die Währung muss genau 3 Grossbuchstaben sein (z.B. CHF)"},
    "err_period_start":    {"EN": "Period start must be YYYY-MM-DD",   "DE": "Periodenstart muss im Format JJJJ-MM-TT sein"},
    "err_period_end":      {"EN": "Period end must be YYYY-MM-DD",     "DE": "Periodenende muss im Format JJJJ-MM-TT sein"},
    "err_company":         {"EN": "Company name is required",          "DE": "Firmenname ist erforderlich"},
    "err_tin":             {"EN": "TIN is required",                   "DE": "UID/TIN ist erforderlich"},
    "err_tin_issued":      {"EN": "TIN issued by must be exactly 2 uppercase letters (e.g. CH)",
                            "DE": "TIN ausgestellt von muss genau 2 Grossbuchstaben sein (z.B. CH)"},
    "err_rec_jur":         {"EN": "Partner country (RecJurCode) must be exactly 2 uppercase letters (e.g. DE)",
                            "DE": "Partnerstaat (RecJurCode) muss genau 2 Grossbuchstaben sein (z.B. DE)"},
    "upe_rules_label":     {"EN": "UPE GloBE Rules",                  "DE": "UPE GloBE-Regeln"},
    "upe_rules_help":      {"EN": "GloBE rules applicable to the Ultimate Parent Entity (GIR201–GIR205). Use GIR204 for QDMTT filers.",
                            "DE": "Anwendbare GloBE-Regeln für die oberste Muttergesellschaft (GIR201–GIR205). GIR204 für QDMTT-Einreicher."},
    "upe_globe_status_label": {"EN": "UPE GloBE Status",            "DE": "UPE GloBE-Status"},
    "upe_globe_status_help":  {"EN": "GloBE entity classification of the Ultimate Parent Entity (GIR301–GIR318). GIR301 for standard Constituent Entity.",
                               "DE": "GloBE-Entitätsstatus der obersten Muttergesellschaft (GIR301–GIR318). GIR301 für normale Untereinheit."},
    "in_coop":             {"EN": "in cooperation with",             "DE": "in Zusammenarbeit mit"},
    "summary_label":       {"EN": "Plain Language Summary",            "DE": "Verständliche Zusammenfassung"},
    "sum_filing":          {"EN": "Filing Information",               "DE": "Einreichungsangaben"},
    "sum_company":         {"EN": "Company",                          "DE": "Unternehmen"},
    "sum_tin":             {"EN": "TIN",                              "DE": "UID/TIN"},
    "sum_jurisdiction":    {"EN": "Jurisdiction",                     "DE": "Jurisdiktion"},
    "sum_period":          {"EN": "Reporting Period",                 "DE": "Berichtsperiode"},
    "sum_currency":        {"EN": "Currency",                         "DE": "Währung"},
    "sum_fas":             {"EN": "Accounting Standard",              "DE": "Rechnungslegungsstandard"},
    "sum_role":            {"EN": "Filing Role",                      "DE": "Einreichungsrolle"},
    "sum_partner":         {"EN": "Partner Country",                  "DE": "Partnerstaat"},
    "sum_income":          {"EN": "GloBE Income Computation",        "DE": "GloBE-Einkommensberechnung"},
    "sum_fanil":           {"EN": "Adjusted FANIL",                   "DE": "Angepasstes FANIL"},
    "sum_income_adj":      {"EN": "Income Adjustments (non-zero)",   "DE": "Einkommensanpassungen (ungleich null)"},
    "sum_net_income":      {"EN": "Net GloBE Income",                "DE": "Netto GloBE-Einkommen"},
    "sum_tax":             {"EN": "Covered Tax Computation",         "DE": "Berechnung der anrechenbaren Steuer"},
    "sum_curr_tax":        {"EN": "Aggregate Current Tax",           "DE": "Aggregierte laufende Steuer"},
    "sum_tax_adj":         {"EN": "Tax Adjustments (non-zero)",      "DE": "Steueranpassungen (ungleich null)"},
    "sum_adj_tax":         {"EN": "Adjusted Covered Tax",            "DE": "Angepasste anrechenbare Steuer"},
    "sum_result":          {"EN": "Result",                           "DE": "Ergebnis"},
    "sum_etr":             {"EN": "Effective Tax Rate (ETR)",        "DE": "Effektiver Steuersatz (ETR)"},
    "disclaimer_label":    {"EN": "Disclaimer",                        "DE": "Haftungsausschluss"},
    "disclaimer_text":     {
        "EN": (
            "This tool is provided for informational purposes only and does not constitute legal or tax advice. "
            "The generated XML output should be reviewed and validated by a qualified tax professional before "
            "submission to the ESTV. MME accepts no liability for errors, inaccuracies, or omissions in the "
            "output, or for any consequences arising from its use. Always verify the final file against the "
            "official ESTV XSD schema prior to submission."
        ),
        "DE": (
            "Dieses Tool dient ausschliesslich zu Informationszwecken und stellt keine Rechts- oder "
            "Steuerberatung dar. Der generierte XML-Output ist vor der Einreichung bei der ESTV von einer "
            "qualifizierten Fachperson zu prüfen und zu validieren. MME übernimmt keine Haftung für Fehler, "
            "Ungenauigkeiten oder Auslassungen im Output oder für Folgen, die sich aus dessen Verwendung "
            "ergeben. Die finale Datei ist vor der Einreichung stets gegen das offizielle ESTV XSD-Schema "
            "zu validieren."
        ),
    },
}


# ─── STREAMLIT UI ────────────────────────────────────────────────────────────

MME_LOGO_PATH  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mme_logo.svg")
ESTV_PEM_PATH   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "estv-publickey.pem")
MUTARA_LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mutara_logo.png")

def _load_mutara_logo_b64() -> str | None:
    try:
        with open(MUTARA_LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except OSError:
        return None

_MUTARA_LOGO_B64 = _load_mutara_logo_b64()

def _load_estv_pem() -> bytes | None:
    try:
        with open(ESTV_PEM_PATH, "rb") as f:
            return f.read()
    except OSError:
        return None

_BUNDLED_PEM = _load_estv_pem()

def _load_mme_logo_svg():
    try:
        with open(MME_LOGO_PATH, "r") as f:
            return f.read()
    except OSError:
        return None

_MME_LOGO_SVG = _load_mme_logo_svg()

_favicon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "favicon.png")
_favicon = Image.open(_favicon_path) if os.path.exists(_favicon_path) else "🌐"

st.set_page_config(
    page_title="GloBE XML Export | MME",
    page_icon=_favicon,
    layout="centered",
)

# ── Language state ────────────────────────────────────────────────────────────
if "lang" not in st.session_state:
    st.session_state["lang"] = "EN"
lang = st.session_state["lang"]

st.markdown("""
<style>
    html, body, [class*="css"] {
        font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
    }

    /* Hide the radio label */
    div[data-testid="stRadio"] label { display: none; }

    /* Style language radio as a pill toggle */
    div[data-testid="stRadio"] > div {
        display: flex; gap: 4px; justify-content: flex-end;
    }
    div[data-testid="stRadio"] > div > label {
        display: flex !important;
        padding: 2px 10px;
        border: 1px solid #e2eaf1;
        border-radius: 20px;
        font-size: 0.78rem;
        font-weight: 600;
        color: #6a7681;
        cursor: pointer;
    }
    div[data-testid="stRadio"] > div > label[data-checked="true"] {
        background: #e0303c;
        border-color: #e0303c;
        color: white;
    }

    [data-testid="stHeader"] { background-color: #FFFFFF; }
    [data-testid="stSidebar"] { background-color: #f5f9fc; }

    .stButton > button[kind="primary"] {
        background-color: #e0303c !important;
        border: none !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 30px !important;
        padding: 0.4rem 1.5rem !important;
    }
    .stButton > button[kind="primary"]:hover { background-color: #c0272f !important; }

    .stDownloadButton > button {
        background-color: #e0303c !important;
        border: none !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 30px !important;
        width: 100%;
    }
    .stDownloadButton > button:hover { background-color: #c0272f !important; }

    h2 { color: #313c45 !important; border-bottom: 2px solid #e0303c; padding-bottom: 4px; }
    [data-testid="stMetricValue"] { color: #e0303c !important; font-weight: 700; font-size: 1rem !important; }
    [data-testid="stMetricLabel"] { font-size: 0.75rem !important; }
    hr { border-color: #e2eaf1 !important; }
    .stCaption { color: #6a7681 !important; }
</style>
""", unsafe_allow_html=True)

# ── Header row: logo + title | language toggle ────────────────────────────────
if _MME_LOGO_SVG:
    scaled_svg = _MME_LOGO_SVG.replace(
        'width="204" height="65"',
        'width="90" height="29" viewBox="0 0 204 65"',
    )
else:
    scaled_svg = ""

hdr_left, hdr_right = st.columns([5, 1])
with hdr_left:
    st.markdown(
        f"""<div style='display:flex; align-items:center; gap:24px; padding:12px 0 8px 0;'>
            <div style='display:flex; flex-direction:column; align-items:flex-start;'>
                <div>{scaled_svg}</div>
                <span style='font-size:0.68rem; color:#6a7681; letter-spacing:0.03em; margin-top:3px;'>v{VERSION}</span>
            </div>
            <h1 style='margin:0; color:#313c45; font-size:1.4rem; font-weight:700;
                font-family:"Helvetica Neue",Helvetica,Arial,sans-serif;'>
                GloBE Information Return (GIR)
            </h1>
        </div>""",
        unsafe_allow_html=True,
    )
with hdr_right:
    selected = st.radio(
        "lang", ["EN", "DE"],
        index=0 if lang == "EN" else 1,
        horizontal=True,
        label_visibility="collapsed",
        key="lang_radio",
    )
    if selected != lang:
        st.session_state["lang"] = selected
        st.rerun()

if _MUTARA_LOGO_B64:
    st.markdown(
        f"""<div style='display:flex; align-items:center; gap:10px; padding:6px 0 10px 0;'>
            <span style='font-size:0.78rem; color:#6a7681;'>{T["in_coop"][lang]}</span>
            <img src='data:image/png;base64,{_MUTARA_LOGO_B64}'
                 style='height:18px; width:auto; opacity:0.85;'>
        </div>""",
        unsafe_allow_html=True,
    )

st.divider()

# ── Step 1: Upload ────────────────────────────────────────────────────────────
st.header(T["step1"][lang])
uploaded = st.file_uploader(
    T["upload_label"][lang],
    type=["xlsx", "xlsm"],
    help=T["upload_help"][lang],
)

# ── Step 2: Company details ───────────────────────────────────────────────────
st.header(T["step2"][lang])

col1, col2 = st.columns(2)
with col1:
    company_name  = st.text_input(T["company_name"][lang], value="PLACEHOLDER_COMPANY_AG")
    tin_value     = st.text_input(T["tin"][lang], value="CHE-123456789")
    _tin_sel      = st.selectbox(T["tin_issued_by"][lang], COUNTRY_DISPLAY,
                                 index=_country_idx("CH"))
    tin_issued_by = _tin_sel[:2]
    _jur_sel      = st.selectbox(T["jurisdiction"][lang], COUNTRY_DISPLAY,
                                 index=_country_idx("CH"))
    jurisdiction  = _jur_sel[:2]

with col2:
    currency         = st.selectbox(T["currency"][lang], CURRENCIES,
                                    index=CURRENCIES.index("CHF"))
    fas              = st.selectbox(T["fas_label"][lang], FAS_OPTIONS)
    _period_start_dt = st.date_input(T["period_start"][lang], value=date(2024, 1, 1))
    _period_end_dt   = st.date_input(T["period_end"][lang],   value=date(2024, 12, 31))
    period_start     = _period_start_dt.strftime("%Y-%m-%d")
    period_end       = _period_end_dt.strftime("%Y-%m-%d")

_rec_sel     = st.selectbox(
    T["partner_country"][lang],
    COUNTRY_DISPLAY,
    index=_country_idx("DE"),
    help=T["partner_help"][lang],
)
rec_jur_code = _rec_sel[:2]

with st.expander(T["advanced"][lang]):
    reporting_role = st.selectbox(
        T["filing_role"][lang],
        ["GIR401", "GIR402", "GIR404"],
        format_func=lambda x: {
            "GIR401": T["gir401"][lang],
            "GIR402": T["gir402"][lang],
            "GIR404": T["gir404"][lang],
        }[x],
        help=T["filing_role_help"][lang],
    )
    tin_type = st.selectbox(
        T["tin_type_label"][lang],
        ["GIR3001", "GIR3002"],
        format_func=lambda x: {
            "GIR3001": T["gir3001"][lang],
            "GIR3002": T["gir3002"][lang],
        }[x],
        help=T["tin_type_help"][lang],
    )
    cfs_of_upe = st.selectbox(
        T["cfs_upe"][lang],
        ["GIR501", "GIR502", "GIR503"],
        format_func=lambda x: {
            "GIR501": T["gir501"][lang],
            "GIR502": T["gir502"][lang],
            "GIR503": T["gir503"][lang],
        }[x],
        help=T["cfs_upe_help"][lang],
    )
    upe_rules = st.selectbox(
        T["upe_rules_label"][lang],
        UPE_RULES_OPTIONS,
        index=UPE_RULES_OPTIONS.index("GIR204"),
        format_func=lambda x: UPE_RULES_LABELS[x],
        help=T["upe_rules_help"][lang],
    )
    upe_globe_status = st.selectbox(
        T["upe_globe_status_label"][lang],
        UPE_GLOBE_STATUS_OPTIONS,
        index=UPE_GLOBE_STATUS_OPTIONS.index("GIR301"),
        format_func=lambda x: UPE_GLOBE_STATUS_LABELS[x],
        help=T["upe_globe_status_help"][lang],
    )
    submission_mode = st.radio(
        T["submission_mode"][lang],
        ["production", "test"],
        format_func=lambda x: T["mode_production"][lang] if x == "production" else T["mode_test"][lang],
        index=1,
        help=T["mode_help"][lang],
        horizontal=True,
    )

# ── Step 3: Export ────────────────────────────────────────────────────────────
st.header(T["step3"][lang])
st.write("")

def validate_inputs(cfg: dict) -> list[str]:
    L = st.session_state.get("lang", "EN")
    errors = []
    if not re.match(r"^[A-Z]{2}$", cfg["jurisdiction"]):
        errors.append(T["err_jurisdiction"][L])
    if not re.match(r"^[A-Z]{3}$", cfg["currency"]):
        errors.append(T["err_currency"][L])
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", cfg["period_start"]):
        errors.append(T["err_period_start"][L])
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", cfg["period_end"]):
        errors.append(T["err_period_end"][L])
    if not cfg["company_name"].strip():
        errors.append(T["err_company"][L])
    if not cfg["tin_value"].strip():
        errors.append(T["err_tin"][L])
    if not re.match(r"^[A-Z]{2}$", cfg["tin_issued_by"]):
        errors.append(T["err_tin_issued"][L])
    if not re.match(r"^[A-Z]{2}$", cfg["rec_jur_code"]):
        errors.append(T["err_rec_jur"][L])
    return errors


if st.button(T["generate_btn"][lang], type="primary", disabled=uploaded is None):
    if uploaded is None:
        st.error(T["upload_first"][lang])
    else:
        with st.spinner(T["spinner_gen"][lang]):
            try:
                file_bytes = uploaded.read()
                data = read_excel(file_bytes)

                cfg = {
                    "company_name":    company_name,
                    "tin_value":       tin_value,
                    "tin_issued_by":   tin_issued_by,
                    "tin_type":        tin_type,
                    "reporting_role":  reporting_role,
                    "rec_jur_code":    rec_jur_code.strip().upper(),
                    "currency":        currency,
                    "jurisdiction":    jurisdiction,
                    "fas":             fas,
                    "cfs_of_upe":      cfs_of_upe,
                    "period_start":    period_start,
                    "period_end":      period_end,
                    "upe_rules":       upe_rules,
                    "upe_globe_status": upe_globe_status,
                }

                input_errors = validate_inputs(cfg)
                if input_errors:
                    for err in input_errors:
                        st.error(err)
                    st.stop()

                xml_str = build_xml(data, cfg, test_mode=(submission_mode == "test"))
                st.session_state["xml_str"]      = xml_str
                st.session_state["xml_filename"] = f"gir_{period_end[:4]}_{jurisdiction}.xml"

                etr_val = fmt_etr(data["adjusted_cov_tax"], data["net_globe_income"])
                cols = st.columns(4)
                cols[0].metric("AdjustedFANIL",  f"{data['adjusted_fanil']:,}")
                cols[1].metric("NetGlobeIncome",  f"{data['net_globe_income']:,}")
                cols[2].metric("AdjustedCovTax",  f"{data['adjusted_cov_tax']:,}")
                cols[3].metric("ETR",              etr_val)

                checks  = validate_xml(xml_str)
                n_pass  = sum(1 for _, ok, _ in checks)
                n_total = len(checks)
                all_ok  = n_pass == n_total

                with st.expander(
                    f"{'✅' if all_ok else '⚠️'}  {T['validation_title'][lang]} — "
                    f"{n_pass}/{n_total} {T['checks_passed'][lang]}",
                    expanded=not all_ok,
                ):
                    for label, ok, detail in checks:
                        icon = "✅" if ok else "❌"
                        if detail and not ok:
                            st.markdown(f"{icon} &nbsp; **{label}**  \n"
                                        f"&nbsp;&nbsp;&nbsp;&nbsp;`{detail}`")
                        else:
                            st.markdown(f"{icon} &nbsp; {label}")
                    if not all_ok:
                        st.caption(T["fix_issues"][lang])

                if all_ok:
                    st.success(T["all_passed"][lang])

                filename = f"gir_{period_end[:4]}_{jurisdiction}.xml"
                st.download_button(
                    label=T["download_xml"][lang],
                    data=xml_str.encode("utf-8"),
                    file_name=filename,
                    mime="application/xml",
                )

                # Plain language summary
                with st.expander(T["summary_label"][lang]):
                    C = cfg["currency"]

                    def _row(label, value, bold=False):
                        ca, cb = st.columns([2, 3])
                        ca.markdown(
                            f"<span style='color:#6a7681;font-size:0.84rem;'>{label}</span>",
                            unsafe_allow_html=True,
                        )
                        color = "#e0303c" if bold else "#313c45"
                        weight = "700" if bold else "600"
                        cb.markdown(
                            f"<span style='color:{color};font-size:0.84rem;font-weight:{weight};'>{value}</span>",
                            unsafe_allow_html=True,
                        )

                    st.markdown(f"**{T['sum_filing'][lang]}**")
                    _row(T["sum_company"][lang],     cfg["company_name"])
                    _row(T["sum_tin"][lang],         cfg["tin_value"])
                    _row(T["sum_jurisdiction"][lang],cfg["jurisdiction"])
                    _row(T["sum_period"][lang],      f"{cfg['period_start']} – {cfg['period_end']}")
                    _row(T["sum_currency"][lang],    cfg["currency"])
                    _row(T["sum_fas"][lang],         cfg["fas"])
                    _row(T["sum_role"][lang],        cfg["reporting_role"])
                    _row(T["sum_partner"][lang],     cfg["rec_jur_code"])

                    st.markdown("---")
                    st.markdown(f"**{T['sum_income'][lang]}**")
                    _row(T["sum_fanil"][lang], f"{data['adjusted_fanil']:,} {C}")

                    nz_income = [(code, amt) for code, amt in data["income_adj"] if amt != 0]
                    if nz_income:
                        st.caption(T["sum_income_adj"][lang])
                        for code, amt in nz_income:
                            desc = GIR_INCOME_LABELS.get(code, code)
                            ca, cb = st.columns([2, 3])
                            ca.markdown(
                                f"<span style='color:#6a7681;font-size:0.78rem;'>{code} — {desc}</span>",
                                unsafe_allow_html=True,
                            )
                            cb.markdown(
                                f"<span style='font-size:0.78rem;'>{amt:,} {C}</span>",
                                unsafe_allow_html=True,
                            )

                    _row(T["sum_net_income"][lang], f"{data['net_globe_income']:,} {C}", bold=True)

                    st.markdown("---")
                    st.markdown(f"**{T['sum_tax'][lang]}**")
                    _row(T["sum_curr_tax"][lang], f"{data['aggregate_curr_tax']:,} {C}")

                    nz_tax = [(code, amt) for code, amt in data["cov_tax_adj"] if amt != 0]
                    if nz_tax:
                        st.caption(T["sum_tax_adj"][lang])
                        for code, amt in nz_tax:
                            desc = GIR_TAX_LABELS.get(code, code)
                            ca, cb = st.columns([2, 3])
                            ca.markdown(
                                f"<span style='color:#6a7681;font-size:0.78rem;'>{code} — {desc}</span>",
                                unsafe_allow_html=True,
                            )
                            cb.markdown(
                                f"<span style='font-size:0.78rem;'>{amt:,} {C}</span>",
                                unsafe_allow_html=True,
                            )

                    _row(T["sum_adj_tax"][lang], f"{data['adjusted_cov_tax']:,} {C}", bold=True)

                    st.markdown("---")
                    st.markdown(f"**{T['sum_result'][lang]}**")
                    etr_pct = float(etr_val) * 100
                    _row(T["sum_etr"][lang], f"{etr_val} ({etr_pct:.2f}%)", bold=True)

                with st.expander(T["preview_xml"][lang]):
                    st.code(xml_str, language="xml")

            except KeyError as e:
                logging.exception("Sheet not found during Excel read")
                st.error(T["sheet_not_found"][lang].format(e))
            except Exception as e:
                logging.exception("XML generation failed")
                st.error(T["error_msg"][lang].format(e))

elif uploaded is None:
    st.info(T["upload_to_enable"][lang])

# ── Step 4: Encrypt for ESTV ──────────────────────────────────────────────────
st.divider()
st.header(T["step4"][lang])

xml_ready = "xml_str" in st.session_state

# Determine which PEM to use
pem_override = None

st.markdown(
    f"<div style='background:#f5f9fc; border-left:3px solid #e0303c; padding:10px 14px;"
    f"border-radius:4px; margin-bottom:12px; font-size:0.88rem; color:#313c45;'>"
    f"🔑 &nbsp;{T['bundled_key_info'][lang]}</div>",
    unsafe_allow_html=True,
)

with st.expander(T["override_key"][lang]):
    pem_file = st.file_uploader(T["pem_label"][lang], type=["pem", "cer"])
    if pem_file is not None:
        pem_override = pem_file.read()
        st.success(T["override_active"][lang])

pem_bytes_to_use = pem_override if pem_override else _BUNDLED_PEM

if st.button(
    T["encrypt_btn"][lang],
    type="primary",
    disabled=(not xml_ready or pem_bytes_to_use is None),
):
    if not xml_ready:
        st.error(T["generate_first"][lang])
    elif pem_bytes_to_use is None:
        st.error(T["upload_pem"][lang])
    else:
        with st.spinner(T["spinner_enc"][lang]):
            try:
                zip_bytes    = encrypt_for_estv(st.session_state["xml_str"], pem_bytes_to_use)
                base_name    = st.session_state["xml_filename"].replace(".xml", "")
                zip_filename = f"{base_name}_encrypted.zip"
                st.download_button(
                    label=T["download_zip"][lang],
                    data=zip_bytes,
                    file_name=zip_filename,
                    mime="application/zip",
                )
                st.success(T["ready_estv"][lang])
            except Exception as e:
                logging.exception("Encryption failed")
                st.error(T["enc_failed"][lang].format(e))

if not xml_ready:
    st.info(T["gen_first_long"][lang])

st.divider()
with st.expander(T["disclaimer_label"][lang]):
    st.markdown(
        f"<p style='color:#6a7681; font-size:0.85rem;'>{T['disclaimer_text'][lang]}</p>",
        unsafe_allow_html=True,
    )
st.markdown(
    "<p style='color:#6a7681; font-size:0.75rem; margin:0;'>MME Legal | Tax | Compliance</p>",
    unsafe_allow_html=True,
)
