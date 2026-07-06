# GIR Correction Mode (Rektifikat) — Pre-Flight Plan

> ## STATUS UPDATE 2026-07-06 — Phase A BUILT & VERIFIED (same day)
> - **Streamlit v2.7.0, commit `b729248` — LOCAL ONLY, not pushed** (push = Streamlit-Cloud deploy; Daniel decides).
> - Design deviation from §2.2 (improvement): no `build_xml(correction=…)` param — `apply_correction()`
>   **post-processes** a freshly built Neumeldung. `build_xml`/`validate_xml` untouched → parity by construction.
> - **Tests 21/21 green** (`tests/test_correction.py`): parse (real PureFert 28-section file + Nord Stream shape),
>   identical-regeneration → all-unchanged, single-change detection, golden correction test+prod flavor, Storno,
>   merged view, 8 negative rule tests, Neumeldung parity guard, XSD validation of the correction file.
> - **UI verified** (Playwright E2E with real files) + Daniel produced a correct correction file
>   (`01-Input/301 Test_gir_2024_CH_korrektur.xml`: GIR102, FilingInfo OECD10 resend, 2× OECD12, chains OK).
> - **Nord Stream (Firma 2) live-proven** on the real production XML `01-Input/101 gir_2024_CH.xml`
>   (6 computations + 5 SH Summaries, PRODUCTION): parse, per-jurisdiction correction, Layer 1 15/15.
>   ⚠️ The submitted figures came from the **FF Excel** — a real correction must start from the SAME Excel
>   that produced the submission, and **verify `101`'s MessageRefId (`CH2024CH513e158f-…`) against the
>   client's acceptance status message** to prove `101` is the accepted original.
> - **Finding — status messages (D2 fallback):** an **accepted** status message contains **NO DocRefIds**
>   (verified: `V2  Test/2.3.3 status-message-….xml` — only `OriginalMessageRefID`). Rejected ones carry
>   `DocRefIDInError` + `FieldPath`, but only for errored records. → No-original recovery ladder: (1) raw-XML
>   download from whoever generated it, (2) ePortal re-download (unverified), (3) ask ESTV support for the
>   DocRefIds citing the MessageRefId. Without DocRefIds, no correction AND no Storno is possible.
> - **ESTV test portal is DOWN** → Phase B (test-portal proof) pending; Phase C additionally needs Toni's original XML.

**Date:** 2026-07-06 · **Status:** Phase A DONE (see box above) — plan below kept as written
**Trigger:** Toni (MME) needs a Rektifikat for a client with an accepted first GIR filing (email 2026-07-04).
**Requirement (Daniel):** must work for Firma 1 (single-jurisdiction, HAS test portal) AND Firma 2
(multi-jurisdiction, NO test portal → correction goes straight to production). Every correction rule
becomes a `validate_xml` check. Prove end-to-end on Firma 1's test portal BEFORE Firma 2 goes productive.

**Sources verified for this plan** (not from memory):
- ESTV `Technische-Wegleitung-GIR-de.md` Kap. 5.3.4 (RecJurCode), 5.3.5 (DocSpec), 6.1–6.4 (Meldesequenzen)
- `globe-xsd/oecdglobetypes_v5.0.xsd` — `DocSpec_Type` has optional `CorrDocRefId` (line 241); schema supports corrections as-is, **no XSD change needed**
- `globe_xml_app_v2.py` (v2.6.3) / `gir_core.py` (portal v3.3.1, parity-locked) — current hardcodes confirmed:
  `MessageTypeIndic = "GIR101"` (app line 612), `doc_type_indic = OECD11/OECD1` (623–625),
  per-section `add_docspec()` with random UUID DocRefIds (627–631)

---

## 0. Decisions (recommended — confirm before build)

| # | Decision | Recommendation | Why |
|---|----------|----------------|-----|
| D1 | Target app | **Streamlit tool first** (`globe_xml_app_v2.py`), mirror into portal core later (Stream-B pattern) | Both production filings were made with the Streamlit tool; MME staff use it today; fastest path to the real Rektifikat. Portal stays parity-locked until proven. |
| D2 | UX approach | **Upload original XML** (option a) + corrected Excel; app extracts DocRefIds and auto-detects changed sections. Manual DocRefId override as fallback (for the recover-from-ESTV-status-message case). | Manual UUID entry (option b) is an error factory — one typo = production bounce for Firma 2. |
| D3 | Storno (OECD3) | Ship the validation rules, expose in UI only as an "Advanced" per-section option | Toni's case is a correction, not a deletion; rules cost nothing, UI stays simple. |
| D4 | Working "original" for the build | Generate fresh Neumeldungen from the two proven test Excels and keep the XMLs (Daniel's plan) | Sufficient for build + test-portal proof. **The REAL original XML/ZIP from Toni stays mandatory for the production Rektifikat** — regenerated files have different UUIDs than what ESTV has on record. |

UTPRAttribution: our app never emits it → parser/validators treat it like the other section types, no UI needed.

---

## 1. The rule set (the spec — each row becomes a validate_xml check)

"Berichtselement" = FilingInfo, GeneralSection, Summary, JurisdictionSection, UTPRAttribution.
Each carries its own `DocSpec` (DocTypeIndic, DocRefId, optional CorrDocRefId).

### Message level
| Rule | ESTV code |
|------|-----------|
| Correction message: `MessageTypeIndic = "GIR102"` | Kap. 6.3.2 |
| New, unique `MessageRefId` — never reuse any earlier one, incl. the original's | 60001 / 6.3.2 |
| No mixing: GIR102 message may contain ONLY OECD2/OECD3 sections (+ the OECD0/OECD10 FilingInfo resend); GIR101 may contain only new (OECD1/OECD11) | 60004 |

### FilingInfo (special: must be in EVERY correction message)
| Rule | ESTV code |
|------|-----------|
| Unchanged FilingInfo → resend: `DocTypeIndic = OECD0` (prod) / `OECD10` (test), **SAME DocRefId as the last transmitted FilingInfo**, NO CorrDocRefId | 60014, 60012 |
| Resend only valid if that DocRefId exists in an earlier message and was never corrected/deleted (≠ any earlier CorrDocRefId) | 60014 |
| FilingInfo itself corrected → OECD2, new DocRefId, CorrDocRefId → previous FilingInfo DocRefId; all subsequent messages must then reference/resend the NEW FilingInfo | Kap. 6.4.3 |
| FilingInfo deletable only if all other sections are deleted too | 60010 |

### Corrected/deleted sections (GeneralSection, Summary, JurisdictionSection, UTPRAttribution)
| Rule | ESTV code |
|------|-----------|
| Replaced **as a whole** (even for a one-field fix); unchanged sections are **OMITTED**, never resent | Kap. 6.3.1 |
| `DocTypeIndic = OECD2` correction / `OECD3` Storno (test: OECD12 / OECD13) | 6.3.2 |
| New unique DocRefId per corrected section — no reuse of ANY earlier DocRefId, incl. the corrected one's | 60007 |
| DocRefId format `CH` + Berichtssteuerperiode + unique id (regex `CH[0-9]{4}.{1,194}`); period must match MessageRefId's | 60011 |
| `CorrDocRefId` REQUIRED on OECD2/OECD3; FORBIDDEN on OECD0/OECD1 | 60015 / 60012 |
| CorrDocRefId must equal the DocRefId of a previously **accepted** element of the **same type** | 60008 / 60005 |
| Each element correctable only once → chains: 2nd correction points at the 1st correction's DocRefId, never the initial one; every CorrDocRefId used exactly once (also unique within the message) | 6.3.3 / 60009 / 60006 |
| Storno ends the chain — a deleted element can only come back as a NEW element in a GIR101 Neumeldung (new DocRefId, no CorrDocRefId, RecJurCode may change) | 6.4.3 |
| `RecJurCode` of a correction must EQUAL the original's; to change it: Storno + Neumeldung | 98200 |
| Test/prod consistency: prod file (no `Test_` filename) must not contain OECD10–13 | 50009 |

### Cross-element rules still apply to the MERGED state
98201 (GeneralSection must carry all RecJurCodes used anywhere) and the value-comparison rules
(60022, 70008, 70036–70053, 70099, 70100) are evaluated by ESTV against the report **as it stands after
the correction** (original sections + replacements). A correction of one section can therefore force
corrections of dependent sections. → Design consequence: we validate a **merged view**, see §2.4.

---

## 2. Design — what gets built (Streamlit tool, then portal mirror)

### 2.1 New module: original-XML parser (`parse_original_xml`)
Input: the originally submitted XML (or the `_encrypted.zip`? **No** — the zip is ESTV-public-key
encrypted, we cannot decrypt it; require the **plain XML** the app offered as "Download raw XML".
The UI must say this explicitly.)
Output (the "correction context"):
```
{ message_ref_id, reporting_period, test_mode,
  filing_info: {doc_ref_id},
  sections: [ {type: GeneralSection|Summary|JurisdictionSection|UTPRAttribution,
               jurisdiction,            # from JurisdictionName / n-th position
               rec_jur_code, doc_ref_id, raw_element} ] }
```
Keyed by **(type, jurisdiction)** — that's the join key to the regenerated file. Multi-jurisdiction
(Firma 2) simply yields more entries; identical logic (per-section DocSpec confirmed in
`gir_core.py:641`).

**Correction chains:** the user must upload the **latest accepted** file. If a correction was already
submitted before, the latest correction XML carries the DocRefIds to reference. UI hint + manual
per-section DocRefId override for the recovery case (DocRefIds read from an ESTV status message).

### 2.2 `build_xml` extension
New optional param `correction: dict | None` (default None → behavior byte-identical to today; parity
tests must still pass). When set:
- `MessageTypeIndic = "GIR102"`; fresh MessageRefId (existing format `CH{year}CH{uuid}` already unique)
- FilingInfo: DocTypeIndic `OECD0`/`OECD10`, DocRefId **copied from the original**, no CorrDocRefId
  (v1 scope: FilingInfo correction itself = out of scope for the UI, rules still validated)
- Emit ONLY the sections flagged changed, each with DocTypeIndic `OECD2`/`OECD12` (or OECD3/OECD13
  Storno), fresh DocRefId, `CorrDocRefId = ` that section's original DocRefId
- `RecJurCode` copied from the original section (not recomputed)

### 2.3 Section diff (auto-detect what changed)
Regenerate the full XML from the corrected Excel (existing pipeline, untouched) → canonicalize each
section (drop `DocSpec`, ElementTree canonical serialization) → byte-compare against the original's
sections by (type, jurisdiction). Result pre-populates a review table:
| Section | Jurisdiction | Status (unchanged / CHANGED / new / missing) | include? |
- **CHANGED** → pre-checked as OECD2
- **new** (in Excel but not in original) → hard warning: not allowed in a GIR102 message; needs a
  separate GIR101 Neumeldung with FilingInfo OECD0-resend (Wegleitung 6.4.2). v1: block + explain.
- **missing** (in original but not in Excel) → offer Storno (OECD3) behind the Advanced toggle.
- User can override any checkbox (tax-expert judgment beats the diff).

### 2.4 `validate_xml` — correction checks (all gating, active only in correction mode)
New signature: `validate_xml(xml_str, correction_ctx=None)`. Two layers:

**Layer 1 — the correction file itself** (every §1 rule):
C1 MessageTypeIndic=GIR102 · C2 MessageRefId ≠ original's + format · C3 no-mixing (all sections
OECD2/3, FilingInfo OECD0; no OECD1) · C4 FilingInfo DocRefId == original's, no CorrDocRefId ·
C5 every section has CorrDocRefId (60015) · C6 every CorrDocRefId resolves to an original DocRefId of
the same type+jurisdiction (60008/60005/98200-adjacent) · C7 CorrDocRefIds unique within message
(60006) · C8 new DocRefIds unique + disjoint from ALL original DocRefIds (60007) + format/period
(60011) · C9 RecJurCode per corrected section == original's (98200) · C10 test/prod DocTypeIndic
consistency (50009, incl. `Test_` filename rule at download) · C11 XSD validation (existing check,
must pass with CorrDocRefId present).

**Layer 2 — the merged view:** replace the corrected sections inside the original XML in memory, then
run the EXISTING structural checks (98201, 70012, 70030, TIN checks, Summary/Jurisdiction consistency
…) against that merged document. This is how we catch "you corrected the DE JurisdictionSection but
the Summary now contradicts it" **before** ESTV does — critical for Firma 2 (no test portal).

Download gate unchanged: encrypt/download disabled until ALL checks green.

### 2.5 UI (Streamlit, step structure preserved)
- Step 1 gains a mode radio: **Neumeldung** (default, today's flow untouched) / **Korrektur (Rektifikat)**
- Correction mode: two uploads (original XML + corrected Excel) → section review table (§2.3) →
  existing Steps 2–4 (company details / Validate file / Download for ESTV)
- Filename: test corrections get the `Test_` prefix automatically (50009)
- EN/DE strings for everything (existing TXT dict pattern)

### 2.6 Out of scope for v1 (explicitly)
FilingInfo content corrections (UI), adding new sections via GIR101-resend flow (6.4.2), decrypting
`_encrypted.zip`, portal mirror (separate step after test-portal proof), gir-form.

---

## 3. Test plan — prove it before Firma 2 touches production

### Phase A — local golden tests (no portal)
1. Regenerate Neumeldungen from both proven Excels (Firma 1 single-jur, Firma 2 multi-jur/Nord-Stream
   structure) and keep the XMLs as fixtures ("originals").
2. Round-trip: `parse_original_xml` on both fixtures → DocRefId count/keys correct (≈4 sections vs
   full multi-jur set).
3. Golden correction: change ONE figure in the Firma-2 Excel (e.g. DE computation) → correction file
   contains exactly FilingInfo(OECD0, original DocRefId) + DE JurisdictionSection(OECD2, CorrDocRefId
   = original DE DocRefId) + (if the change ripples) the DE/affected Summary — and NOTHING else;
   XSD-valid; all Layer-1+2 checks green.
4. Negative tests — each check must FAIL on purpose-built bad files: reused DocRefId, missing
   CorrDocRefId, CorrDocRefId pointing at wrong type, changed RecJurCode, OECD1 mixed in, wrong
   FilingInfo DocRefId, test-indic in prod file.
5. Parity guard: `correction=None` output byte-identical to v2.6.3 (existing parity test style).

### Phase B — Firma 1 test portal, end-to-end (the proof)
1. Submit fresh **OECD11 Neumeldung** (test) → accepted → keep XML.
2. Build **OECD12 correction** against it in the app → submit → **accepted = feature proven**.
3. Chain proof: second correction pointing at the FIRST correction's DocRefIds → accepted.
   (Also implicitly proves 60009: pointing at the initial element again would bounce.)
4. Save all status messages to `V2  Test/` as evidence.

### Phase C — the real Rektifikat (Firma 2 / Toni's client, production)
**Gate: Phase B fully green AND the real original XML (or its DocRefIds) from Toni in hand.**
Load real original + corrected Excel → review table with the tax expert → all checks green →
encrypt → production submit → archive status message.

---

## 4. Risk register
| Risk | Mitigation |
|------|-----------|
| Client can't find the original raw XML | Fallback: DocRefIds from the ESTV portal status message via manual override fields; verify portal lets him re-download the submission (unverified) |
| Original was produced by an older app version → section canonicalization diff shows false CHANGED | Diff is advisory only; user confirms the section set. Layer-1 checks don't depend on the diff |
| Hidden ESTV correction rules not in the Wegleitung (like 98201 was) | Phase B surfaces them on the test portal at zero cost; that's why B gates C |
| Ripple effects (70008ff) after correcting one section | Layer-2 merged-view validation catches locally |
| Firma 1 test portal access lapsed | Confirm with Toni BEFORE building Phase B assumptions |

## 5. Rollout & effort
Build order: parser → build_xml correction param → Layer-1 checks → diff+UI → Layer-2 merged
validation → tests. Version **v2.7.0** (XML-engine change), push → Streamlit Cloud auto-deploy
(sandbox push caveat: `dangerouslyDisableSandbox`, Daniel decides when). Portal mirror (gir_core.py
+ FastAPI routes) as its own follow-up with parity tests, after Phase B.

**Open with Toni (blocking Phase C only):** original XML/ZIP or DocRefIds; confirm test portal access;
confirm which sections are wrong (drives the correction content).
