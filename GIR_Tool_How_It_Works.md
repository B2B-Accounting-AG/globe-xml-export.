# GloBE Information Return (GIR) Tool — How It Works

A plain-language guide to what the tool does with your data, and the decisions it makes
automatically. It is a **data-conversion tool**: it turns your filled-in GIR Excel template into
the official OECD GIR XML and encrypts it for the ESTV portal. It does **not** calculate or judge
any tax position — the output reflects only what you enter.

---

## 1. What goes in, what comes out

| | |
|---|---|
| **Input** | Your completed **GIR Excel template** (the OECD/ESTV multi-sheet workbook). |
| **Output** | (1) the GIR **XML** file, and (2) the **encrypted .zip** ready to upload to the ESTV GIR portal. |

The tool reads three kinds of sheet from the workbook:

| Sheet | Used for |
|-------|----------|
| `1 MNE Group Information` | Filing entity, the group, every constituent entity, and the ownership structure. |
| `3 GloBE Computations` | The full ETR / top-up-tax computation for the **computed jurisdiction** (Switzerland). |
| `2 Safe Harbours XX` (one per country) | Each jurisdiction's **safe-harbour election** and supporting figures. |

---

## 2. The steps

1. **Upload** the Excel file.
2. The tool **auto-fills** the company details (name, TIN, jurisdiction, currency, accounting standard,
   reporting period) from sheet 1 — you can review and edit any of them.
3. **Review the constituent entities** (and enter any missing TINs) and the **safe-harbour jurisdictions**.
4. **Generate** the XML. The tool runs structural checks and shows any problems.
5. **Encrypt** the XML for ESTV.
6. **Submit** to the ESTV **test (ABN) portal** first, read the status message, then file for real.

---

## 3. The decisions the tool makes (the "if → then" logic)

### A. Which jurisdictions get reported
- **If** a `Safe Harbours XX` tab has a safe-harbour elected → **then** that jurisdiction is reported
  with a *Summary* + a *safe-harbour section*.
- **If** a `Safe Harbours XX` tab is **empty** (no election) → **then** it is **ignored**.
- **If** a jurisdiction has the full computation in `3 GloBE Computations` (Switzerland) → **then** it
  gets a *full computation section* (and, if it also elects a safe harbour, a Summary as well).

### B. How each safe harbour is reported
| Safe harbour elected (`Safe Harbours` tab) | How the tool reports it |
|---|---|
| `GIR1202` QDMTT safe harbour | Summary (with ETR range, SBIE, QDMTT top-up tax) — the computation comes from sheet 3. |
| `GIR1203/1204/1205` Transitional CbCR (de-minimis / ETR / routine-profit tests) | Safe-harbour section with the CbCR figures (Revenue, Profit, Income Tax) + a sub-group marker. |
| `GIR1206` Transitional UTPR safe harbour | Safe-harbour section with the corporate income-tax rate. |

### C. Entity TINs (tax numbers)
- **If** a constituent entity has a TIN in the template → **then** it is used.
- **If** a constituent entity's TIN is **missing** → **then** the structural check **fails (red)** and
  the entity is listed. ESTV requires a real TIN for **every** entity, so the file cannot be accepted
  until each one is filled in (in the in-app table or in sheet 1, row "5. TIN").
- Each entity needs **its own** local tax number — they do **not** share the parent company's TIN.

### D. Ownership (who owns each entity)
- **If** the template says an entity is owned by the **parent (UPE)** → **then** the parent's TIN is
  used as the owner.
- **If** an entity is owned by **another group company** *and* that owner's TIN is provided in sheet 1
  (row "9. TIN") → **then** that owner TIN is used.
- **If** an entity is owned by another group company **but the owner's TIN is blank** → **then** the
  tool reports it as **parent-owned** (a valid simplification ESTV accepts). To show the true
  intermediate ownership instead, fill the owner's TIN in sheet 1 row 9.

### E. Test vs. real filing
- **Test mode** (default) tags the file as a **test** submission for the ESTV ABN/acceptance portal.
- **Production mode** is for the real filing on the live portal. Switch to it only when you actually file.

---

## 4. What the tool checks (and what it does **not**)

**It checks:** that the XML is well-formed, matches the official GIR schema, has the required fields,
and that every entity has a TIN. It will not let an obviously incomplete file look "ready".

**It does not check:** whether your **numbers are correct**, whether the **right safe harbour** was
elected, or any **tax position**. Those are your responsibility. The only authoritative validation is
the ESTV portal's own status message — **always test-submit first.**

---

## 5. What you need to provide for a real filing

1. **A real TIN for every constituent entity** (not a placeholder).
2. **Optional:** the direct-owner TIN (sheet 1, row 9) for entities owned by another group company, if
   you want the true ownership chain reflected.
3. **Switch to Production mode** when filing for real.

---

*This document explains the tool's behaviour for orientation only. It is not legal or tax advice; have
the return reviewed by a qualified professional before filing.*
