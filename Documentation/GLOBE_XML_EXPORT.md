# GloBE XML Export — Documentation

**MME Legal | Tax | Compliance** — in cooperation with Mutara  
Swiss QDMTT 2024 · OECD GIR XML Schema (January 2025) · **v1.5.5**

---

## Overview

This tool converts the Swiss QDMTT (Qualified Domestic Minimum Top-up Tax) Excel calculation template into a valid OECD **GloBE Information Return (GIR) XML** file, following the OECD Pillar Two XML Schema published in January 2025.

Two delivery formats are available:

| Format | File | Who uses it |
|---|---|---|
| Web app (Streamlit) | `globe_xml_app.py` | Any browser — no installation required |
| Excel macro (VBA) | `ExportGlobEXML.bas` | Windows users (no Python needed) |

**Web app features:**
- Bilingual interface (EN / DE toggle)
- Built-in structural validation (20+ checks)
- Plain language summary of key figures
- One-click encryption for ESTV — no separate encryptor tool needed
- Bundled ESTV public key (valid until 2027-02-04)
- Test / Production submission mode toggle (DocTypeIndic OECD10 vs OECD1)
- GeneralSection with CorporateStructure/UPE block (required by ESTV since v1.5.0)
- Zero-value adjustment filtering (prevents rule 70060)

---

## Files

```
App/
├── globe_xml_app.py          # Streamlit web app (main entry point)
├── convert_to_globe_xml.py   # Command-line Python script
├── ExportGlobEXML.bas        # VBA macro (import into .xlsm)
├── requirements.txt          # Python dependencies
├── config.toml               # App theme (MME brand colours)
├── VERSION                   # Current version number
├── mme_logo.svg              # MME logo
├── mutara_logo.png           # Mutara logo
├── favicon.png               # Browser tab icon (MME three-bar mark)
├── estv-publickey.pem        # Bundled ESTV public key (DigiCert, valid until 2027-02-04)
└── output/
    └── gir_YYYY_CH.xml       # Generated XML files
```

---

## Prerequisites

### Web app (Streamlit Cloud — no installation needed)
The app runs on Streamlit Community Cloud. Users access it via browser with no local setup required.

### Running locally
```bash
pip3 install streamlit openpyxl cryptography Pillow
streamlit run App/globe_xml_app.py
```
Then open **http://localhost:8501**.

### Command-line script
```bash
pip3 install openpyxl
```

### VBA macro
No installation required. Import `ExportGlobEXML.bas` into the `.xlsm` file once (see [VBA Setup](#vba-macro-setup)).

---

## Using the Web App

### Start

The app is deployed on **Streamlit Community Cloud** (workspace `b2b-accounting-ag`, repo `globe-xml-export`, branch `main`).

A **language toggle (EN / DE)** is available top-right and persists throughout the session.

### Step 1 — Upload Excel file
Upload the completed `Calculation File.xlsx` or `.xlsm`.  
The file must contain a sheet named exactly **`QDMTT 2024`**.

### Step 2 — Company details

| Field | Input type | Example / Default | Notes |
|---|---|---|---|
| Company name | Text | `Muster AG` | Legal entity name |
| TIN | Text | `CHE-123456789` | Swiss UID number |
| TIN issued by | Dropdown | `CH – Switzerland` | ISO 3166-1 Alpha-2 country list |
| Jurisdiction | Dropdown | `CH – Switzerland` | ISO 3166-1 Alpha-2 country list |
| Currency | Dropdown | `CHF` | ISO 4217 list |
| Financial Accounting Standard | Dropdown | `Swiss GAAP FER` | Swiss GAAP FER, IFRS, US GAAP, UK GAAP, HGB, Local GAAP |
| Period start | Date picker | `2024-01-01` | Calendar selector |
| Period end | Date picker | `2024-12-31` | Calendar selector |
| Partner country (RecJurCode) | Dropdown | `CH – Switzerland` | Receiving jurisdiction; must differ from Jurisdiction |

**Advanced options** (defaults are correct for Swiss QDMTT):

| Field | Default | Options |
|---|---|---|
| Filing role | `GIR401 — Ultimate Parent Entity (UPE)` | GIR401 = UPE, GIR402 = DFE, GIR404 = CE (matches ESTV ePortal roles) |
| TIN type | `GIR3001 — Tax Identification Number (TIN)` | GIR3001 = TIN, GIR3002 = Functional equivalent |
| CFS of UPE | `GIR501 — Consolidated Financial Statement (subparagraph a)` | GIR501–GIR503 |
| Submission mode | `Test / CTS (OECD11)` | **Test / CTS** for `eportal-a.admin.ch` (OECD11 = new submission in test mode); **Production (OECD1)** for `eportal.admin.ch` |

### Step 3 — Export
Click **Generate XML**. The app will:
1. Read the jurisdictional totals from Column N of the Excel sheet
2. Build the GIR XML structure
3. Run 20 structural validation checks
4. Show key metrics (FANIL, NetGlobeIncome, AdjustedCovTax, ETR)
5. Offer the XML file for download

A **Plain Language Summary** expander shows a human-readable breakdown of filing info, GloBE income adjustments (non-zero only), covered tax adjustments, and ETR — in the selected language.

### Step 4 — Encrypt for ESTV
Click **Encrypt & Download**. No key upload required — the ESTV public key is bundled in the app.

| Detail | Value |
|---|---|
| Bundled key | `estv-publickey.pem` (DigiCert / encryptor.estv.admin.ch) |
| Valid until | 2027-02-04 |
| Override | Expand "Use a different key" to upload a custom `.pem` or `.cer` |

The app produces an encrypted `.zip` ready to upload directly to the ESTV GIR-Applikation — **no separate ESTV Encryptor tool needed**:

| File in ZIP | Contents |
|---|---|
| `Payload` | `Payload.xml` compressed (ZIP DEFLATE) then AES-256-CBC encrypted |
| `Key` | AES key + IV (48 bytes) RSA PKCS#1 v1.5 encrypted with ESTV public key |

> - Generate the XML in Step 3 first — the Encrypt button is disabled until XML has been generated in the current session.
> - Max upload size on the ePortal: **10 MB**.

---

## Using the Command-Line Script

Edit the `CONFIG` block at the top of `convert_to_globe_xml.py`, then run:

```bash
python3 convert_to_globe_xml.py
```

Output is written to `output/gir_2024_CH.xml`.

---

## VBA Macro Setup (Windows)

**One-time setup per workbook:**

1. Save the Excel file as `.xlsm` (macro-enabled)
2. Open the VBA editor: `Alt + F11`
3. `File → Import File` → select `ExportGlobEXML.bas`
4. Update the constants at the top of the module:
   ```vba
   Private Const COMPANY_NAME As String = "Muster AG"
   Private Const TIN_VALUE    As String = "CHE-123456789"
   ```
5. Run `AddExportButton` once to add the export button to the sheet
6. Close the VBA editor

**Exporting:**  
Click the **Export to GloBE XML** button on the sheet. The file is saved to `output\gir_2024_CH.xml` in the same folder as the workbook.

> If a macro security warning appears on open: click **Enable Content**. This is a one-time prompt per workbook.

---

## Excel Template Structure

The tool reads from sheet **`QDMTT 2024`** only.

**Column N** contains jurisdictional totals (`=SUM(F:M)` across up to 8 entity columns).

### Summary rows (Column N)

| Row | XML Element |
|---|---|
| 236 | `FANIL` / `AdjustedFANIL` |
| 264 | `NetGlobeIncome / Total` |
| 295 | `IncomeTaxExpense` / `AggregrateCurrentTax` |
| 314 | `AdjustedCoveredTax / Total` |

### NetGlobeIncome adjustments (Column N, rows 238–263)

| Row | GIR Code | Article |
|---|---|---|
| 238 | GIR2001 | Net Taxes Expense – Art 3.2.1(a) |
| 239 | GIR2002 | Excluded Dividends – Art 3.2.1(b) |
| 240 | GIR2003 | Excluded Equity Gain/Loss – Art 3.2.1(c) |
| 241 | GIR2004 | Included Revaluation Method Gain/Loss – Art 3.2.1(d) |
| 242 | GIR2005 | Gain/loss on disposition excluded – Art 3.2.1(e) |
| 243 | GIR2006 | Asymmetric FX Gains/Losses – Art 3.2.1(f) |
| 244 | GIR2007 | Policy Disallowed Expenses – Art 3.2.1(g) |
| 245 | GIR2008 | Prior Period Errors – Art 3.2.1(h) |
| 246 | GIR2009 | Changes in Accounting Principles – Art 3.2.1(h) |
| 247 | GIR2010 | Accrued Pension Expense – Art 3.2.1(i) |
| 248 | GIR2011 | Debt releases – Art 3.2.1 |
| 249 | GIR2012 | Stock-based compensation – Art 3.2.2 |
| 250 | GIR2013 | Arm's length adjustments – Art 3.2.3 |
| 251 | GIR2014 | Qualified Refundable Tax Credit / MTTC – Art 3.2.4 |
| 252 | GIR2015 | Election: Gains/losses using realisation principle – Art 3.2.5 |
| 253 | GIR2016 | Election: Adjusted Asset Gain – Art 3.2.6 |
| 254 | GIR2017 | Intragroup Financing Arrangement expense – Art 3.2.7 |
| 255 | GIR2018 | Election: intragroup transactions same jurisdiction – Art 3.2.8 |
| 256 | GIR2019 | Insurance company taxes charged to policyholders – Art 3.2.9 |
| 257 | GIR2020 | Additional Tier One Capital – Art 3.2.10 |
| 258 | GIR2021 | CE joining/leaving MNE Group – Art 3.2.11 & 6.2 |
| 259 | GIR2022 | Reduction (Flow-through Entity UPE) – Art 3.2.11 & 7.1 |
| 260 | GIR2023 | Reduction (Deductible Dividend Regime UPE) – Art 3.2.11 & 7.2 |
| 261 | GIR2024 | Taxable Distribution Method election – Art 3.2.11 & 7.6 |
| 262 | GIR2025 | International Shipping Income – Art 3.3 |
| 263 | GIR2026 | Transactions between CEs – Art 9.1.3 |

### AdjustedCoveredTax adjustments (Column N, rows 297–313 + Column H rows 95–96)

> Note: The Excel template uses internal GIR24xx codes. The XML schema requires GIR27xx.

| Row | Col | XML GIR Code | Article |
|---|---|---|---|
| 297 | N | GIR2701 | Covered Tax accrued as expense – Art 4.1.2(a) |
| 298 | N | GIR2703 | Covered Taxes – uncertain tax position – Art 4.1.2(c) |
| 299 | N | GIR2704 | Qualified RFTC/MTTC reduction – Art 4.1.2(d) |
| 300 | N | GIR2705 | Qualified Flow-through Tax Benefits – Art 3.2.1(c) |
| 301 | N | GIR2706 | Current tax on excluded income – Art 4.1.3(a) |
| 302 | N | GIR2707 | Non-Qualified credits / Other credits – Art 4.1.3(b) |
| 303 | N | GIR2708 | Covered Taxes refunded/credited – Art 4.1.3(c) |
| 304 | N | GIR2709 | Current tax – uncertain tax position – Art 4.1.3(d) |
| 305 | N | GIR2710 | Current tax not paid within 3 years – Art 4.1.3(e) |
| 306 | N | GIR2711 | Post-filing adjustments – Art 4.6.1 |
| 307 | N | GIR2712 | Covered Taxes – Net Asset Gain/Loss – Art 3.2.6 |
| 308 | N | GIR2713 | Reduction (Flow-through Entity UPE) – Art 7.1 |
| 309 | N | GIR2714 | Covered Taxes – Deductible Dividend Regime – Art 7.2.2 |
| 310 | N | GIR2715 | Deemed Distribution Tax – Art 7.3 |
| 311 | N | GIR2716 | Taxable Distribution Method – Art 7.6.2(b) |
| 312 | N | GIR2717 | Total Deferred Tax Adjustment Amount – Art 4.4.1(b) |
| 313 | N | GIR2718 | Increase/decrease in equity/OCI – Art 4.1.1(c) |
| 95 | H | GIR2719 | Excess Neg Tax Expense generated – Art 4.1.5 & 5.2.1 |
| 96 | H | GIR2720 | Excess Neg Tax Expense utilised – Art 4.1.5 & 5.2.1 |

> GIR2702 (GloBE Loss DTA – Art 4.5) has no row in this template. Add manually if applicable.

---

## Calculated Values

| XML Element | Calculation |
|---|---|
| `ETRRate` | `AdjustedCoveredTax ÷ NetGlobeIncome` (clamped 0–1, 4 decimal places) |
| `TopUpTaxPercentage` | Fixed `0.0000` (no top-up tax for QDMTT-qualified entities) |
| `MessageRefId` | `{jurisdiction}{year}{jurisdiction}{uuid}` |
| `Timestamp` | UTC time of generation |

---

## Structural Validation Checks

The app runs 20 checks automatically after every export:

| # | Check |
|---|---|
| 1 | Well-formed XML |
| 2 | Root element (`globe:GLOBE_OECD`) present with `globe:MessageSpec` child |
| 3 | MessageSpec — all required fields present (incl. `SendingEntityIN`) |
| 4 | `MessageRefId` format (`CH[year]CH[uuid]`) |
| 5 | Timestamp format (`YYYY-MM-DDTHH:MM:SS`) |
| 6 | Period dates format (`YYYY-MM-DD`) |
| 7 | Company name — not placeholder |
| 8 | FilingCE Role (GIR401–GIR405) |
| 9 | TIN — not placeholder |
| 10 | TIN attributes (`issuedBy` + `TypeOfTIN`) |
| 11 | FilingInfo DocSpec (`DocTypeIndic` + `DocRefId`) |
| 12 | JurisdictionSection `RecJurCode` present |
| 13 | JurisdictionSection DocSpec (`DocTypeIndic` + `DocRefId`) |
| 14 | Currency `currCode` attribute |
| 15 | OverallComputation — all required elements present |
| 16 | ETRRate format (decimal `0.0000`–`1.0000`) |
| 17 | TopUpTaxPercentage format |
| 18 | NetGlobeIncome adjustment codes are valid GIR codes (zero-value items filtered) |
| 19 | AdjustedCoveredTax adjustment codes are valid GIR codes (zero-value items filtered) |
| 20 | All monetary amounts are integers (no decimals) |
| 21 | GeneralSection present with CorporateStructure / UPE block |

---

## Output XML Structure

```
globe:GLOBE_OECD (xmlns:globe="urn:oecd:ties:globe:v2" version="1.0")
├── globe:MessageSpec
│   ├── globe:SendingEntityIN      TIN of filing entity  ← must be first
│   ├── globe:TransmittingCountry
│   ├── globe:ReceivingCountry
│   ├── globe:MessageType          GIR
│   ├── globe:MessageRefId
│   ├── globe:MessageTypeIndic     GIR101
│   ├── globe:ReportingPeriod
│   └── globe:Timestamp
└── globe:GLOBEBody
    ├── globe:FilingInfo
    │   ├── globe:FilingCE         ResCountryCode, Name, TIN, Role
    │   ├── globe:AccountingInfo   CFSofUPE, FAS, Currency
    │   ├── globe:Period           Start, End
    │   ├── globe:NameMNE
    │   └── globe:DocSpec          stf:DocTypeIndic (OECD1 / OECD10), stf:DocRefId
    ├── globe:GeneralSection                          ← added v1.5.0
    │   ├── globe:RecJurCode       Same as jurisdiction (CH)
    │   ├── globe:CorporateStructure
    │   │   └── globe:UPE
    │   │       └── globe:OtherUPE
    │   │           └── globe:ID   Name, ResCountryCode, TIN, Rules, GlobeStatus
    │   └── globe:DocSpec          stf:DocTypeIndic (OECD1 / OECD10), stf:DocRefId
    └── globe:JurisdictionSection
        ├── globe:RecJurCode       Partner/receiving jurisdiction
        ├── globe:Jurisdiction
        ├── globe:GLoBETax / ETR / ETRStatus / ETRComputation / OverallComputation
        │   ├── globe:FANIL
        │   ├── globe:AdjustedFANIL
        │   ├── globe:NetGlobeIncome
        │   │   ├── globe:Total
        │   │   └── globe:Adjustments × n   (GIR2001–GIR2026, non-zero only)
        │   ├── globe:IncomeTaxExpense
        │   ├── globe:ETRRate
        │   ├── globe:TopUpTaxPercentage
        │   ├── globe:AdjustedCoveredTax
        │   │   ├── globe:Total
        │   │   ├── globe:AggregrateCurrentTax
        │   │   └── globe:Adjustments × n   (GIR2701–GIR2720, non-zero only)
        │   ├── globe:ExcessProfits
        │   ├── globe:QDMTT
        │   ├── globe:TopUpTax
        │   └── globe:ExcessNegTaxExpense
        └── globe:DocSpec          stf:DocTypeIndic (OECD1 / OECD10), stf:DocRefId
```

**Namespaces used:**

| Prefix | URI | Used for |
|---|---|---|
| `globe` | `urn:oecd:ties:globe:v2` | All GIR elements (root + body) |
| `stf` | `urn:oecd:ties:globestf:v5` | DocSpec children (`DocTypeIndic`, `DocRefId`) |
| `xsi` | `http://www.w3.org/2001/XMLSchema-instance` | `schemaLocation` attribute |

---

## Test Environment

ESTV provides a dedicated test environment at `https://eportal-a.admin.ch/`.

| | |
|---|---|
| Test window | 7 April – 3 July 2026 |
| Invitation code | Not sent by post — email `gir-test@estv.admin.ch` with your ESTV-ID (052.XXXX.XXXX) and registration date |
| Behaviour | Test submissions are processed normally; you receive a status response |
| DocTypeIndic | Use `OECD10` (Test / CTS mode in Advanced Options) |

**ESTV error codes encountered:**

| Code | Meaning | Status |
|---|---|---|
| 50007 | Schema validation failed — root element or namespace not recognised | Fixed in v1.4.1 |
| 50008 | `DocTypeIndic` outside accepted range — OECD10 on production portal, or OECD1/mixed on CTS | Triggered by v1.5.1 and v1.5.2 (see Known Issues) |
| 50009 | Production file contains test DocTypeIndic (OECD10–13) or filename starts with "Test" | Avoid by using correct submission mode |
| 60013 | OECD0 (= OECD10 in test) used in non-FilingInfo `DocTypeIndic` — OECD10 is Resend and only valid for FilingInfo | Fixed in v1.5.3 — use OECD11 (new submission) in GeneralSection and JurisdictionSection |
| 60014 | Unknown/invalid DocRefId for resend — fired because OECD10 (Resend) requires an existing DocRefId | Fixed in v1.5.4 — use OECD11 everywhere for new submissions |
| 60022 | `GIR401` FilingCE TIN does not match any TIN in UPE element | Fixed in v1.5.0 (GeneralSection added) |
| 70060 | `GIR2025` present but `IntShippingIncome` element missing | Fixed in v1.5.0 (zero-value filter) |
| 98201 | GeneralSection missing or does not contain all RecJurCodes | Fixed in v1.5.0 (GeneralSection added) |

---

## DocTypeIndic Reference

Confirmed by ESTV (Tobias Buser, 2026-05-22). Test and production codes must **never be mixed** in the same message.

| Production | Test | Meaning | Valid in |
|---|---|---|---|
| `OECD1` | `OECD11` | New submission (Neumeldung) | All DocSpec positions |
| `OECD2` | `OECD12` | Correction | All DocSpec positions except FilingInfo |
| `OECD3` | `OECD13` | Deletion | All DocSpec positions except FilingInfo |
| `OECD0` | `OECD10` | Resend | **FilingInfo only** — only after a prior accepted submission |

> FilingInfo is a mandatory element and must always be included, even when only other sections change. On resubmissions, FilingInfo uses `OECD0` / `OECD10` (Resend); other sections use the appropriate correction/deletion code.

---

## CTS Test Results (PureFert Holding AG, 2026-05-21/22)

| Version | DocTypeIndic used | ValidationResult | Notes |
|---|---|---|---|
| v1.5.0 | OECD10 everywhere | Rejected | 60013, 60014 — OECD10 is Resend, not valid for new submission in non-FilingInfo positions |
| v1.5.1 | OECD1 everywhere | Rejected | 50008 — production codes not accepted on CTS portal |
| v1.5.2 | OECD10 (FilingInfo) + OECD1 (sections) | Rejected | 50008 — mixed production/test codes |
| v1.5.3 | OECD10 (FilingInfo) + OECD11 (sections) | Rejected | 60014 — OECD10 (Resend) still wrong for FilingInfo on first submission |
| v1.5.4 | OECD11 everywhere | **Accepted** ✓ | Correct — OECD11 = new submission in test mode |

---

## Final Submission

Once all 20 structural checks pass:

1. Validate locally against the OECD XSD (available in `Documentation/globe-xsd/GLOBEXML_v1.0.xsd`):
   ```bash
   python3 -c "
   from lxml import etree
   schema = etree.XMLSchema(etree.parse('Documentation/globe-xsd/GLOBEXML_v1.0.xsd'))
   doc = etree.parse('App/output/gir_2024_CH.xml')
   print('Valid:', schema.validate(doc))
   print(schema.error_log)
   "
   ```

2. In the app, set **Submission mode** in Advanced Options:
   - `Test / CTS (OECD11)` → upload to `https://eportal-a.admin.ch/`
   - `Production (OECD1)` → upload to `https://eportal.admin.ch/`

3. Click **Encrypt & Download** — the encrypted ZIP is produced immediately using the bundled ESTV public key. No separate encryptor tool required.

4. Upload the encrypted ZIP to the **myESTV portal → GIR-Applikation** (max 10 MB).

## Disclaimer

This tool is provided for informational purposes only and does not constitute legal or tax advice. The generated XML output should be reviewed and validated by a qualified tax professional before submission to the ESTV. MME accepts no liability for errors, inaccuracies, or omissions in the output, or for any consequences arising from its use.

---

## Contact

Eidgenössische Steuerverwaltung  
Abteilung Informationsaustausch in Steuersachen — Team AIA  
Eigerstrasse 65, 3003 Bern  
Email: `info-gir@estv.admin.ch`  
Tel: +41 58 466 78 76

---

## Reference

- OECD GloBE Information Return XML Schema User Guide, January 2025  
  DOI: [10.1787/c594935a-en](https://doi.org/10.1787/c594935a-en)
- OECD GIR XML Schema: `GLOBEXML_v1.0.xsd` (`Documentation/globe-xsd/`)
- Swiss ESTV Technische Wegleitung GIR (`Documentation/Technische-Wegleitung-GIR-de.pdf`)
- XML Namespace (GIR elements): `urn:oecd:ties:globe:v2` with prefix `globe:`
- XML Namespace (DocSpec): `urn:oecd:ties:globestf:v5` with prefix `stf:`
- Swiss QDMTT legal basis: Art. 4 MinBestG (Mindestbesteuerungsgesetz)
