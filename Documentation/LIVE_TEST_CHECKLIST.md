# Live Test Checklist — GloBE XML Export

## Before the Session

### 1. Request Test Portal Access
- Email `gir-test@estv.admin.ch` with your ESTV-ID (`052.XXXX.XXXX`) and registration date
- Test environment: `https://eportal-a.admin.ch/`
- Test window: 7 April – 3 July 2026

### 2. Prepare the Excel File
- Use the completed QDMTT calculation file (`.xlsx` or `.xlsm`)
- Must contain a sheet named exactly **`QDMTT 2024`**

### 3. Install the App (if not already done)
```bash
pip install -r requirements.txt
streamlit run globe_xml_app.py
```
Open `http://localhost:8501` in the browser.

---

## During the Session

### Step 1 — Upload Excel
Upload the QDMTT calculation file.

### Step 2 — Fill in Company Details

| Field | Notes |
|-------|-------|
| Company name | Legal entity name (not the placeholder) |
| TIN | Real Swiss UID, e.g. `CHE-123456789` |
| TIN issued by | `CH – Switzerland` |
| Jurisdiction | `CH – Switzerland` |
| Currency | `CHF` |
| Financial Accounting Standard | e.g. Swiss GAAP FER |
| Period start / end | `2024-01-01` / `2024-12-31` |
| Partner country (RecJurCode) | The receiving jurisdiction — must **not** be CH |

**Advanced options** (check against ESTV ePortal registration):

| Field | Check |
|-------|-------|
| Filing role | UPE / DFE / CE — must match what's registered in ePortal |
| TIN type | GIR3001 (standard TIN) |
| CFS of UPE | GIR501 (Consolidated FS) |
| **Submission mode** | **Test / CTS (OECD10)** for `eportal-a.admin.ch` — switch to Production (OECD1) only when filing to `eportal.admin.ch` |

### Step 3 — Generate XML
- Click **Generate XML**
- All 20 structural validation checks must pass before proceeding

### Step 4 — Encrypt for ESTV
- Click **Encrypt & Download**
- The ESTV public key is **bundled in the app** — no separate key upload required
- Output: `gir_2024_CH_encrypted.zip`
- If you need to override the key: expand "Use a different key" and upload a custom `.pem` or `.cer`

### Step 5 — Upload to Portal
- Log in to `https://eportal-a.admin.ch/` (test) or `https://eportal.admin.ch/` (production)
- Navigate to **GIR-Applikation**
- Upload the encrypted ZIP (max 10 MB)
- You will receive a status response confirming receipt

---

## Status Response Codes

| Code | Meaning | Action |
|------|---------|--------|
| `Accepted` | File accepted | Done |
| 50007 | Schema validation failed — namespace/root element not recognised | Check app version ≥ 1.4.1 |
| 50008 | DocTypeIndic mismatch | Verify Submission mode in Advanced Options matches the portal (OECD10 for test, OECD1 for production) |
