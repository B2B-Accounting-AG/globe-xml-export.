#!/usr/bin/env python3
"""
GloBE XML Export App
Streamlit web UI for converting the Swiss QDMTT Excel template to OECD GIR XML.
Run: streamlit run globe_xml_app.py
"""

import io
import logging
import re
import uuid
import openpyxl
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

import os
import streamlit as st

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)

# ─── XML SETUP ───────────────────────────────────────────────────────────────

GIR_NS = "urn:oecd:ties:gir:v1"
ET.register_namespace("", GIR_NS)
N = f"{{{GIR_NS}}}"


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

ROW_EXCESS_NEG_GENERATED = 95
ROW_EXCESS_NEG_UTILIZED  = 96
EXCESS_NEG_COL = "H"
ROW_ADJUSTED_FANIL     = 236
ROW_NET_GLOBE_INCOME   = 264
ROW_AGGREGATE_CURR_TAX = 295
ROW_ADJUSTED_COV_TAX   = 314


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


def build_xml(data: dict, cfg: dict) -> str:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
    year = cfg["period_end"][:4]
    msg_ref = f"{cfg['jurisdiction']}{year}{cfg['jurisdiction']}{str(uuid.uuid4())}"

    root = ET.Element(N + "GloBE_Message")

    hdr = sub(root, "MessageHeader")
    sub(hdr, "TransmittingCountry", cfg["jurisdiction"])
    sub(hdr, "ReceivingCountry",    cfg["jurisdiction"])
    sub(hdr, "MessageType",         "GIR")
    sub(hdr, "MessageRefID",        msg_ref)
    sub(hdr, "MessageTypeIndic",    "GIR101")
    sub(hdr, "ReportingPeriod",     cfg["period_end"])
    sub(hdr, "Timestamp",           now)
    sub(hdr, "SendingEntityIN",     cfg["tin_value"])

    body = sub(root, "GloBE_Body")
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
    sub(acct, "Currency", currCode=cfg["currency"])

    period = sub(fi, "Period")
    sub(period, "Start", cfg["period_start"])
    sub(period, "End",   cfg["period_end"])
    sub(fi, "NameMNE", cfg["company_name"])

    fi_doc = sub(fi, "DocSpec")
    sub(fi_doc, "DocTypeIndic", "OECD1")
    sub(fi_doc, "DocRefId",     f"{cfg['jurisdiction']}{year}-{str(uuid.uuid4())}")

    jur_sec = sub(body, "JurisdictionSection")
    sub(jur_sec, "RecJurCode", cfg["rec_jur_code"])

    globe_tax  = sub(jur_sec, "GloBE_Tax")
    etr        = sub(globe_tax, "ETR")
    etr_status = sub(etr, "ETR_Status")
    etr_comp   = sub(etr_status, "ETR_Computation")
    oc         = sub(etr_comp, "OverallComputation")

    sub(oc, "FANIL",         data["adjusted_fanil"])
    sub(oc, "AdjustedFANIL", data["adjusted_fanil"])

    ngi = sub(oc, "NetGlobeIncome")
    sub(ngi, "Total", data["net_globe_income"])
    for gir_code, amount in data["income_adj"]:
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
        adj = sub(act, "Adjustments")
        sub(adj, "Amount",         amount)
        sub(adj, "AdjustmentItem", gir_code)

    jur_doc = sub(jur_sec, "DocSpec")
    sub(jur_doc, "DocTypeIndic", "OECD1")
    sub(jur_doc, "DocRefId",     f"{cfg['jurisdiction']}{year}-{str(uuid.uuid4())}")

    ET.indent(root, space="  ")
    buf = io.BytesIO()
    ET.ElementTree(root).write(buf, xml_declaration=True, encoding="utf-8")
    return buf.getvalue().decode("utf-8")


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

    g = {"g": GIR_NS}

    def text(path):
        el = root.find(path, g)
        return el.text.strip() if el is not None and el.text else None

    def findall(path):
        return root.findall(path, g)

    # 2. Namespace
    check("Namespace (urn:oecd:ties:gir:v1)", GIR_NS in root.tag)

    # 3. MessageHeader — all required fields (incl. Swiss SendingEntityIN)
    hdr_fields = ["TransmittingCountry", "ReceivingCountry", "MessageType",
                  "MessageRefID", "MessageTypeIndic", "ReportingPeriod",
                  "Timestamp", "SendingEntityIN"]
    missing_hdr = [f for f in hdr_fields if text(f"g:MessageHeader/g:{f}") is None]
    check("MessageHeader — all required fields (incl. SendingEntityIN)", not missing_hdr,
          f"Missing: {', '.join(missing_hdr)}" if missing_hdr else "")

    # 4. MessageRefID format: CH[0-9]{4}CH...
    msg_ref = text("g:MessageHeader/g:MessageRefID")
    check("MessageRefID format (CH[year]CH[uuid])",
          bool(msg_ref and re.match(r"^[A-Z]{2}\d{4}[A-Z]{2}.+", msg_ref)),
          msg_ref or "missing")

    # 5. Timestamp format
    ts = text("g:MessageHeader/g:Timestamp")
    check("Timestamp format (YYYY-MM-DDTHH:MM:SS)",
          bool(ts and re.match(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}", ts)),
          ts or "missing")

    # 6. Period dates
    start = text("g:GloBE_Body/g:FilingInfo/g:Period/g:Start")
    end   = text("g:GloBE_Body/g:FilingInfo/g:Period/g:End")
    date_ok = bool(
        start and re.match(r"\d{4}-\d{2}-\d{2}$", start) and
        end   and re.match(r"\d{4}-\d{2}-\d{2}$", end)
    )
    check("Period dates (YYYY-MM-DD)", date_ok,
          f"Start: {start}  End: {end}" if not date_ok else "")

    # 7. Company name — not placeholder (FilingCE direct child, no ID wrapper)
    name = text("g:GloBE_Body/g:FilingInfo/g:FilingCE/g:Name")
    name_ok = bool(name and name != "PLACEHOLDER_COMPANY_AG")
    check("Company name (not placeholder)", name_ok,
          "Still set to PLACEHOLDER_COMPANY_AG" if not name_ok else "")

    # 8. Role in FilingCE (GIR401–GIR405)
    role = text("g:GloBE_Body/g:FilingInfo/g:FilingCE/g:Role")
    check("FilingCE Role (GIR401–GIR405)",
          bool(role and re.match(r"^GIR40[1-5]$", role)),
          role or "missing")

    # 9. TIN — not placeholder, has required attributes
    tin_el = root.find("g:GloBE_Body/g:FilingInfo/g:FilingCE/g:TIN", g)
    tin_val = tin_el.text.strip() if tin_el is not None and tin_el.text else None
    tin_ok = bool(tin_val and tin_val != "CHE-123456789")
    check("TIN (not placeholder)", tin_ok,
          "Still set to CHE-123456789" if not tin_ok else "")
    if tin_el is not None:
        check("TIN attributes (issuedBy + TypeOfTIN)",
              bool(tin_el.get("issuedBy") and tin_el.get("TypeOfTIN")))

    # 10. DocSpec in FilingInfo
    fi_doc = root.find("g:GloBE_Body/g:FilingInfo/g:DocSpec", g)
    fi_doc_ok = (
        fi_doc is not None and
        fi_doc.find(f"{N}DocTypeIndic") is not None and
        fi_doc.find(f"{N}DocRefId") is not None
    )
    check("FilingInfo DocSpec (DocTypeIndic + DocRefId)", fi_doc_ok)

    # 11. RecJurCode in JurisdictionSection
    rec_jur = text("g:GloBE_Body/g:JurisdictionSection/g:RecJurCode")
    check("JurisdictionSection RecJurCode present",
          bool(rec_jur and re.match(r"^[A-Z]{2}$", rec_jur)),
          rec_jur or "missing")

    # 12. DocSpec in JurisdictionSection
    jur_doc = root.find("g:GloBE_Body/g:JurisdictionSection/g:DocSpec", g)
    jur_doc_ok = (
        jur_doc is not None and
        jur_doc.find(f"{N}DocTypeIndic") is not None and
        jur_doc.find(f"{N}DocRefId") is not None
    )
    check("JurisdictionSection DocSpec (DocTypeIndic + DocRefId)", jur_doc_ok)

    # 8. Currency currCode
    ccy_el = root.find("g:GloBE_Body/g:FilingInfo/g:AccountingInfo/g:Currency", g)
    check("Currency currCode attribute",
          bool(ccy_el is not None and ccy_el.get("currCode")))

    # 9. OverallComputation — required elements
    oc = ("g:GloBE_Body/g:JurisdictionSection/g:GloBE_Tax/g:ETR"
          "/g:ETR_Status/g:ETR_Computation/g:OverallComputation")
    oc_fields = ["FANIL", "AdjustedFANIL", "IncomeTaxExpense",
                 "ETRRate", "TopUpTaxPercentage"]
    missing_oc = [f for f in oc_fields if text(f"{oc}/g:{f}") is None]
    check("OverallComputation — required elements", not missing_oc,
          f"Missing: {', '.join(missing_oc)}" if missing_oc else "")

    # 10. ETRRate — decimal 0–1, 4 decimal places
    etr_val = text(f"{oc}/g:ETRRate")
    try:
        etr_f  = float(etr_val) if etr_val else None
        etr_ok = (etr_f is not None and 0 <= etr_f <= 1
                  and bool(re.match(r"^\d\.\d{4}$", etr_val)))
    except (ValueError, TypeError):
        etr_ok = False
    check("ETRRate format (0.0000 – 1.0000)", etr_ok, etr_val or "missing")

    # 11. TopUpTaxPercentage format
    tup = text(f"{oc}/g:TopUpTaxPercentage")
    check("TopUpTaxPercentage format (0.0000)",
          bool(tup and re.match(r"^\d\.\d{4}$", tup)), tup or "missing")

    # 12. All 26 NetGlobeIncome adjustment items
    ngi_codes = {el.text for el in findall(
        f"{oc}/g:NetGlobeIncome/g:Adjustments/g:AdjustmentItem") if el.text}
    expected_ngi = {f"GIR{2000+i}" for i in range(1, 27)}
    missing_ngi  = expected_ngi - ngi_codes
    check("NetGlobeIncome — all 26 adjustments (GIR2001–GIR2026)", not missing_ngi,
          f"Missing: {', '.join(sorted(missing_ngi))}" if missing_ngi else "")

    # 13. All 19 AdjustedCoveredTax adjustment items (GIR2701–GIR2720, no GIR2702)
    act_codes = {el.text for el in findall(
        f"{oc}/g:AdjustedCoveredTax/g:Adjustments/g:AdjustmentItem") if el.text}
    expected_act = {f"GIR27{i:02d}" for i in range(1, 21) if i != 2}
    missing_act  = expected_act - act_codes
    check("AdjustedCoveredTax — all 19 adjustments (GIR2701–GIR2720)", not missing_act,
          f"Missing: {', '.join(sorted(missing_act))}" if missing_act else "")

    # 14. All amounts are integers
    non_int = [el.text for el in root.findall(f".//{N}Amount")
               if el.text and "." in el.text]
    check("All amounts are integers (no decimals)", not non_int,
          f"Non-integer: {non_int[:3]}" if non_int else "")

    return results


# ─── STREAMLIT UI ────────────────────────────────────────────────────────────

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "B2B_long.png")

def _load_logo_bytes():
    try:
        with open(LOGO_PATH, "rb") as f:
            return f.read()
    except OSError:
        return None

_LOGO_BYTES = _load_logo_bytes()

st.set_page_config(
    page_title="GloBE XML Export | b2b accounting",
    page_icon="🌐",
    layout="centered",
)

st.markdown("""
<style>
    /* Header bar */
    [data-testid="stHeader"] { background-color: #FFFFFF; }

    /* Sidebar background */
    [data-testid="stSidebar"] { background-color: #F7F4F2; }

    /* Primary button */
    .stButton > button[kind="primary"] {
        background-color: #E05A2B !important;
        border: none !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #C44E24 !important;
    }

    /* Download button */
    .stDownloadButton > button {
        background-color: #E05A2B !important;
        border: none !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
        width: 100%;
    }
    .stDownloadButton > button:hover {
        background-color: #C44E24 !important;
    }

    /* Section headers */
    h2 { color: #5C4F47 !important; border-bottom: 2px solid #E05A2B; padding-bottom: 4px; }

    /* Metric value color */
    [data-testid="stMetricValue"] { color: #E05A2B !important; font-weight: 700; }

    /* Divider */
    hr { border-color: #E8E0DB !important; }

    /* Caption */
    .stCaption { color: #8C7B74 !important; }
</style>
""", unsafe_allow_html=True)

# Header: logo + title
col_logo, col_title = st.columns([1, 3])
with col_logo:
    if _LOGO_BYTES:
        st.image(_LOGO_BYTES, width=160)
with col_title:
    st.markdown(
        "<h1 style='margin-top:18px; color:#3D3330; font-size:1.4rem; font-weight:700;'>"
        "GloBE Information Return<br>"
        "<span style='font-size:0.95rem; color:#8C7B74; font-weight:400;'>"
        "Swiss QDMTT 2024 &nbsp;·&nbsp; OECD GIR XML Schema (January 2025)"
        "</span></h1>",
        unsafe_allow_html=True,
    )

st.divider()

# ── Step 1: Upload ────────────────────────────────────────────────────────────
st.header("1. Upload Excel file")
uploaded = st.file_uploader(
    "Calculation File (.xlsx or .xlsm)",
    type=["xlsx", "xlsm"],
    help='The Swiss QDMTT calculation template with sheet "QDMTT 2024"',
)

# ── Step 2: Company details ───────────────────────────────────────────────────
st.header("2. Company details")

col1, col2 = st.columns(2)
with col1:
    company_name  = st.text_input("Company name", value="PLACEHOLDER_COMPANY_AG")
    tin_value     = st.text_input("TIN", value="CHE-123456789")
    tin_issued_by = st.text_input("TIN issued by (ISO 3166-1 Alpha-2)", value="CH")
    jurisdiction  = st.text_input("Jurisdiction (ISO 3166-1 Alpha-2)", value="CH")

with col2:
    currency     = st.text_input("Currency (ISO 4217)", value="CHF")
    fas          = st.text_input("Financial Accounting Standard", value="Swiss GAAP FER")
    period_start = st.text_input("Period start", value="2024-01-01")
    period_end   = st.text_input("Period end",   value="2024-12-31")

rec_jur_code = st.text_input(
    "Partner country (RecJurCode)",
    value="DE",
    help="ISO 3166-1 Alpha-2 country code of the partner jurisdiction (must not be CH)",
)

with st.expander("Advanced options"):
    reporting_role = st.selectbox(
        "Filing role",
        ["GIR401", "GIR402", "GIR403", "GIR404", "GIR405"],
        format_func=lambda x: {
            "GIR401": "GIR401 — Ultimate Parent Entity (UPE)",
            "GIR402": "GIR402 — Designated Filing Entity (DFE)",
            "GIR403": "GIR403 — Local Filing Entity",
            "GIR404": "GIR404 — Constituent Entity",
            "GIR405": "GIR405 — Other",
        }[x],
        help="Role of the filing constituent entity",
    )
    tin_type   = st.selectbox("TIN type", ["GIR3001", "GIR3002"],
                              help="GIR3001 = TIN, GIR3002 = Functional equivalent")
    cfs_of_upe = st.selectbox("CFS of UPE", ["GIR501", "GIR502", "GIR503"],
                              help="GIR501 = subparagraph a")

# ── Step 3: Export ────────────────────────────────────────────────────────────
st.header("3. Export")

def validate_inputs(cfg: dict) -> list[str]:
    errors = []
    if not re.match(r"^[A-Z]{2}$", cfg["jurisdiction"]):
        errors.append("Jurisdiction must be exactly 2 uppercase letters (e.g. CH)")
    if not re.match(r"^[A-Z]{3}$", cfg["currency"]):
        errors.append("Currency must be exactly 3 uppercase letters (e.g. CHF)")
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", cfg["period_start"]):
        errors.append("Period start must be YYYY-MM-DD")
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", cfg["period_end"]):
        errors.append("Period end must be YYYY-MM-DD")
    if not cfg["company_name"].strip():
        errors.append("Company name is required")
    if not cfg["tin_value"].strip():
        errors.append("TIN is required")
    if not re.match(r"^[A-Z]{2}$", cfg["tin_issued_by"]):
        errors.append("TIN issued by must be exactly 2 uppercase letters (e.g. CH)")
    if not re.match(r"^[A-Z]{2}$", cfg["rec_jur_code"]):
        errors.append("Partner country (RecJurCode) must be exactly 2 uppercase letters (e.g. DE)")
    return errors


if st.button("Generate XML", type="primary", disabled=uploaded is None):
    if uploaded is None:
        st.error("Please upload an Excel file first.")
    else:
        with st.spinner("Reading Excel and building XML…"):
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
                }

                input_errors = validate_inputs(cfg)
                if input_errors:
                    for err in input_errors:
                        st.error(err)
                    st.stop()

                xml_str = build_xml(data, cfg)

                # Summary metrics
                etr_val = fmt_etr(data["adjusted_cov_tax"], data["net_globe_income"])
                cols = st.columns(4)
                cols[0].metric("AdjustedFANIL",  f"{data['adjusted_fanil']:,}")
                cols[1].metric("NetGlobeIncome",  f"{data['net_globe_income']:,}")
                cols[2].metric("AdjustedCovTax",  f"{data['adjusted_cov_tax']:,}")
                cols[3].metric("ETR",              etr_val)

                # Validation
                checks  = validate_xml(xml_str)
                n_pass  = sum(1 for _, ok, _ in checks)
                n_total = len(checks)
                all_ok  = n_pass == n_total

                with st.expander(
                    f"{'✅' if all_ok else '⚠️'}  Structural validation — "
                    f"{n_pass}/{n_total} checks passed",
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
                        st.caption(
                            "Fix the issues above, then re-generate. "
                            "Once all checks pass, validate against the official "
                            "ESTV XSD before submission."
                        )

                if all_ok:
                    st.success("All structural checks passed.")

                # Download
                filename = f"gir_{period_end[:4]}_{jurisdiction}.xml"
                st.download_button(
                    label="⬇️  Download XML",
                    data=xml_str.encode("utf-8"),
                    file_name=filename,
                    mime="application/xml",
                )

                # Preview
                with st.expander("Preview XML"):
                    st.code(xml_str, language="xml")

            except KeyError as e:
                logging.exception("Sheet not found during Excel read")
                st.error(f'Sheet not found: {e}. Make sure the file contains a sheet named "QDMTT 2024".')
            except Exception as e:
                logging.exception("XML generation failed")
                st.error(f"Error: {e}")

elif uploaded is None:
    st.info("Upload the Excel file above to enable export.")

st.divider()
st.markdown(
    "<p style='color:#C8B8B0; font-size:0.75rem; margin:0;'>B2B Accounting AG</p>",
    unsafe_allow_html=True,
)
