# GloBE XML Export — Documentation

**B2B Accounting AG**  
Swiss QDMTT 2024 · OECD GIR XML Schema (January 2025)

---

## Overview

This tool converts the Swiss QDMTT (Qualified Domestic Minimum Top-up Tax) Excel calculation template into a valid OECD **GloBE Information Return (GIR) XML** file, following the OECD Pillar Two XML Schema published in January 2025.

Two delivery formats are available:

| Format | File | Who uses it |
|---|---|---|
| Web app (Streamlit) | `globe_xml_app.py` | Mac / any browser |
| Excel macro (VBA) | `ExportGlobEXML.bas` | Windows employees (no Python needed) |

---

## Files

```
01-Input/
├── globe_xml_app.py          # Streamlit web app
├── convert_to_globe_xml.py   # Command-line Python script
├── ExportGlobEXML.bas         # VBA macro (import into .xlsm)
├── Calculation File.xlsx      # Swiss QDMTT Excel template
├── GLOBE_XML_EXPORT.md        # This document
├── .streamlit/
│   └── config.toml            # App theme (B2B brand colours)
└── output/
    └── gir_YYYY_CH.xml        # Generated XML files
```

---

## Prerequisites

### Web app
```bash
pip3 install streamlit openpyxl
```

### Command-line script
```bash
pip3 install openpyxl
```

### VBA macro
No installation required. Import `ExportGlobEXML.bas` into the `.xlsm` file once (see [VBA Setup](#vba-macro-setup)).

---

## Using the Web App

### Start

The app is deployed on **Streamlit Community Cloud** (`b2b-accounting-ag` workspace, repo `globe-xml-export`, branch `main`).

To run locally:
```bash
streamlit run /Volumes/Claude_Vault/ClaudeCode/03-Clients/MME/globe_xml_app.py
```
Then open **http://localhost:8501**.

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
| Partner country (RecJurCode) | Dropdown | `DE – Germany` | Receiving jurisdiction; must differ from Jurisdiction |

**Advanced options** (defaults are correct for Swiss QDMTT):

| Field | Default | Options |
|---|---|---|
| Filing role | `GIR401 — Ultimate Parent Entity (UPE)` | GIR401–GIR405 |
| TIN type | `GIR3001 — Tax Identification Number (TIN)` | GIR3001 = TIN, GIR3002 = Functional equivalent |
| CFS of UPE | `GIR501 — Consolidated Financial Statement (subparagraph a)` | GIR501–GIR503 |

### Step 3 — Export
Click **Generate XML**. The app will:
1. Read the jurisdictional totals from Column N of the Excel sheet
2. Build the GIR XML structure
3. Run 20 structural validation checks
4. Show key metrics (FANIL, NetGlobeIncome, AdjustedCovTax, ETR)
5. Offer the XML file for download

### Step 4 — Encrypt for ESTV
Upload the **ESTV public key** (`ESTV-PublicKey.pem`) from the myESTV portal, then click **Encrypt & Download**.

The app produces an encrypted `.zip` ready to upload directly to the ESTV GIR-Applikation:

| File in ZIP | Contents |
|---|---|
| `Payload` | `Payload.xml` compressed (ZIP DEFLATE) then AES-256-CBC encrypted |
| `Key` | AES key + IV (48 bytes) RSA PKCS#1 v1.5 encrypted with ESTV public key |

> Generate the XML in Step 3 first — the Encrypt button is disabled until XML has been generated in the current session.

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
| `MessageRefID` | `{jurisdiction}2024{jurisdiction}{random 12-char hex}` |
| `Timestamp` | UTC time of generation |

---

## Structural Validation Checks

The app runs 20 checks automatically after every export:

| # | Check |
|---|---|
| 1 | Well-formed XML |
| 2 | Namespace (`urn:oecd:ties:gir:v1`) |
| 3 | MessageHeader — all required fields present (incl. `SendingEntityIN`) |
| 4 | MessageRefID format (`CH[year]CH[uuid]`) |
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
| 18 | All 26 NetGlobeIncome adjustments present (GIR2001–GIR2026) |
| 19 | All 19 AdjustedCoveredTax adjustments present (GIR2701–GIR2720) |
| 20 | All monetary amounts are integers (no decimals) |

---

## Output XML Structure

```
GloBE_Message (xmlns="urn:oecd:ties:gir:v1")
├── MessageHeader
│   ├── TransmittingCountry
│   ├── ReceivingCountry
│   ├── MessageType            GIR
│   ├── MessageRefID
│   ├── MessageTypeIndic       GIR101
│   ├── ReportingPeriod
│   ├── Timestamp
│   └── SendingEntityIN        TIN of filing entity
└── GloBE_Body
    ├── FilingInfo
    │   ├── FilingCE           ResCountryCode, Name, TIN, Role
    │   ├── AccountingInfo     CFSofUPE, FAS, Currency
    │   ├── Period             Start, End
    │   ├── NameMNE
    │   └── DocSpec            DocTypeIndic, DocRefId
    └── JurisdictionSection
        ├── RecJurCode         Partner/receiving jurisdiction
        ├── GloBE_Tax / ETR / ETR_Status / ETR_Computation / OverallComputation
        │   ├── FANIL
        │   ├── AdjustedFANIL
        │   ├── NetGlobeIncome
        │   │   ├── Total
        │   │   └── Adjustments × 26   (GIR2001–GIR2026)
        │   ├── IncomeTaxExpense
        │   ├── ETRRate
        │   ├── TopUpTaxPercentage
        │   └── AdjustedCoveredTax
        │       ├── Total
        │       ├── AggregrateCurrentTax
        │       └── Adjustments × 19   (GIR2701–GIR2720)
        └── DocSpec            DocTypeIndic, DocRefId
```

---

## Final Submission

Once all 20 structural checks pass:

1. Optionally validate against the official **ESTV XSD** (Swiss Federal Tax Administration):
   ```bash
   xmllint --schema estv_gir.xsd output/gir_2024_CH.xml
   ```
   > The ESTV XSD had not been publicly released as of January 2025. Once available, drop it into the project folder and run the command above.

2. Use **Step 4 — Encrypt for ESTV** in the web app:
   - Upload `ESTV-PublicKey.pem` from the myESTV portal
   - Click **Encrypt & Download** to produce `gir_2024_CH_encrypted.zip`

3. Upload the encrypted ZIP to the **myESTV portal → GIR-Applikation**.

---

## Reference

- OECD GloBE Information Return XML Schema User Guide, January 2025  
  DOI: [10.1787/c594935a-en](https://doi.org/10.1787/c594935a-en)
- XML Namespace: `urn:oecd:ties:gir:v1`
- Swiss QDMTT legal basis: Art. 4 MinBestG (Mindestbesteuerungsgesetz)
