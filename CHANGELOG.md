# Changelog — GloBE XML Export

## v1.4.2 — 2026-05-21
- Add test/production submission mode toggle in Advanced Options
  - `Test / CTS (OECD10)` — default, for `eportal-a.admin.ch`
  - `Production (OECD1)` — for `eportal.admin.ch`
- Fix error 50008: ESTV test portal requires DocTypeIndic `OECD10`, not `OECD1`

## v1.4.1 — 2026-05-21
- Fix `SendingEntityIN` element order — must be first child of `MessageSpec` per XSD sequence constraint
- Fix `MessageRefId` capitalisation in `convert_to_globe_xml.py` (was `MessageRefID`)
- Local XSD validation (`GLOBEXML_v1.0.xsd`) now reports `Valid: True`

## v1.4.0 — 2026-05-21
- Fix XML namespace: `urn:oecd:ties:gir:v1` → `urn:oecd:ties:globe:v2` (correct OECD GIR schema targetNamespace)
- Add STF namespace `urn:oecd:ties:globestf:v5` for DocSpec children
- Fix error 50007: ESTV can now find declaration of element `globe:GLOBE_OECD`
- Add `SendingEntityIN` to MessageSpec (was missing entirely)
- Register namespace prefixes (`globe:`, `stf:`, `xsi:`) for clean XML serialisation
