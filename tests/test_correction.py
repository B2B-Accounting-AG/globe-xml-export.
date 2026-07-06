"""Correction-mode (Rektifikat) tests — GIR102 per ESTV Technische Wegleitung Kap. 5.3.4-6.4.

Loads the pre-UI core of the Streamlit app via exec (the module runs the UI at
import time, so a plain import is not possible). Fixtures:
  - original_firma1_300.xml       real app output (PureFert, 1 computation + 13 SH jurisdictions, OECD11)
  - nordstream_ff_template.xlsx   Firma-2 shape: 6 full computations + 5 SH elections, X5 stateless PEs
  - "V2  Test/20260609_GIR Template.xlsx" (repo)  Firma-1 template

Run: python3 -m pytest tests/test_correction.py -q
"""
import copy
import pathlib
import re
import xml.etree.ElementTree as ET

import pytest

ROOT     = pathlib.Path(__file__).resolve().parent.parent
APP      = ROOT / "App" / "globe_xml_app_v2.py"
FIXTURES = pathlib.Path(__file__).resolve().parent / "fixtures"
XSD      = ROOT / "Documentation" / "globe-xsd" / "GLOBEXML_v1.0.xsd"
PUREFERT = ROOT / "V2  Test" / "20260609_GIR Template.xlsx"
NORDSTREAM = FIXTURES / "nordstream_ff_template.xlsx"

_UI_MARKER = "# ─── STREAMLIT UI"


def _load_core():
    src = APP.read_text()
    assert _UI_MARKER in src, "UI marker moved — update the test loader"
    ns: dict = {"__name__": "gir_app_core_under_test"}
    exec(compile(src.split(_UI_MARKER)[0], str(APP), "exec"), ns)
    return ns


G = _load_core()
N, S = G["N"], G["S"]


def _cfg_from_excel(excel_path: pathlib.Path) -> tuple[list, dict]:
    """Mirror the UI auto-fill: computations + cfg from the template."""
    parsed = G["read_excel_v2"](excel_path.read_bytes())
    m = parsed["mne"]
    # The raw templates ship without CE TINs (ESTV 70006 data gap) — in the app the
    # user completes them in the TIN editor before generating. Mirror that here.
    for i, ce in enumerate(m["constituent_entities"], 1):
        if not ce.get("tin"):
            ce["tin"] = f"TESTTIN-{i:03d}"
    cfg = {
        "company_name":    m["company_name"] or "TESTCO AG",
        "tin_value":       m["tin_value"] or "CHE-999.999.999",
        "tin_issued_by":   m["tin_issued_by"] or "CH",
        "tin_type":        "GIR3001",
        "reporting_role":  m["filing_role"] or "GIR401",
        "rec_jur_code":    "CH",
        "currency":        m["currency"] or "CHF",
        "jurisdiction":    m["jurisdiction"] or "CH",
        "fas":             m["fas"] or "OECD-GloBE",
        "cfs_of_upe":      "GIR501",
        "period_start":    m["period_start"] or "2024-01-01",
        "period_end":      m["period_end"] or "2024-12-31",
        "upe_rules":       m["upe_rules"] or "GIR204",
        "upe_globe_status": m["upe_globe_status"] or "GIR301",
        "constituent_entities": m["constituent_entities"],
        "safe_harbours":        parsed["safe_harbours"],
    }
    return parsed["computations"], cfg


def _sections(xml_str: str):
    body = ET.fromstring(xml_str).find(N + "GLOBEBody")
    return dict(G["_walk_sections"](body))


def _docspec(el):
    ds = el.find(N + "DocSpec")
    get = lambda tag: (ds.find(S + tag).text or "").strip() if ds.find(S + tag) is not None else None
    return get("DocTypeIndic"), get("DocRefId"), get("CorrDocRefId")


# ─── parse_original_xml ───────────────────────────────────────────────────────

def test_parse_firma1_real_file():
    """Daniel's real app output: 1 FilingInfo + 1 GeneralSection + 13 Summaries
    + 13 JurisdictionSections, all OECD11 (test Neumeldung)."""
    ctx = G["parse_original_xml"]((FIXTURES / "original_firma1_300.xml").read_bytes())
    assert ctx["message_type"] == "GIR101"
    assert ctx["reporting_period"] == "2024-12-31"
    assert ctx["filing_doc_ref_id"].startswith("CH2024-")
    assert ctx["test_mode"] is True
    counts: dict = {}
    for s in ctx["sections"]:
        counts[s["type"]] = counts.get(s["type"], 0) + 1
    assert counts == {"GeneralSection": 1, "Summary": 13, "JurisdictionSection": 13}
    # per-section keys unique, all DocRefIds present + distinct
    keys = [s["key"] for s in ctx["sections"]]
    refs = [s["doc_ref_id"] for s in ctx["sections"]]
    assert len(set(keys)) == len(keys)
    assert len(set(refs)) == len(refs)


def test_parse_firma2_multijurisdiction():
    """Firma-2 shape from the Nord Stream FF template: 6 full computations
    (CH/DE/RU/DK/SE/FI) + 5 SH Summaries — parser keys them per jurisdiction."""
    computations, cfg = _cfg_from_excel(NORDSTREAM)
    xml = G["build_xml"](computations, cfg, test_mode=True)
    ctx = G["parse_original_xml"](xml.encode())
    jurs = {s["jurisdiction"] for s in ctx["sections"] if s["type"] == "JurisdictionSection"}
    assert jurs == {"CH", "DE", "RU", "DK", "SE", "FI"}
    n_summaries = sum(1 for s in ctx["sections"] if s["type"] == "Summary")
    assert n_summaries == 5


def test_parse_rejects_non_gir():
    with pytest.raises(ValueError):
        G["parse_original_xml"](b"<foo/>")


# ─── diff ─────────────────────────────────────────────────────────────────────

def test_diff_identical_regeneration_is_all_unchanged():
    """Same Excel + same cfg regenerated → only UUIDs/timestamps differ, and the
    canonical section diff must see NO change."""
    computations, cfg = _cfg_from_excel(NORDSTREAM)
    original = G["build_xml"](computations, cfg, test_mode=True)
    ctx = G["parse_original_xml"](original.encode())
    regenerated = G["build_xml"](computations, cfg, test_mode=True)
    rows = G["diff_correction_sections"](ctx, regenerated)
    assert rows and all(r["status"] == "unchanged" for r in rows)


def test_diff_detects_single_changed_jurisdiction():
    computations, cfg = _cfg_from_excel(NORDSTREAM)
    original = G["build_xml"](computations, cfg, test_mode=True)
    ctx = G["parse_original_xml"](original.encode())
    comp2 = copy.deepcopy(computations)
    de = next(c for c in comp2 if c["jur_iso"] == "DE")
    de["adjusted_fanil"] += 1
    rows = G["diff_correction_sections"](ctx, G["build_xml"](comp2, cfg, test_mode=True))
    changed = [r["key"] for r in rows if r["status"] != "unchanged"]
    assert changed == ["JurisdictionSection|DE"]


# ─── apply_correction: golden path ────────────────────────────────────────────

def _golden_correction(test_mode=True):
    computations, cfg = _cfg_from_excel(NORDSTREAM)
    original = G["build_xml"](computations, cfg, test_mode=test_mode)
    ctx = G["parse_original_xml"](original.encode())
    comp2 = copy.deepcopy(computations)
    next(c for c in comp2 if c["jur_iso"] == "DE")["adjusted_fanil"] += 1
    new_xml = G["build_xml"](comp2, cfg, test_mode=test_mode)
    corr = G["apply_correction"](new_xml, ctx, {"JurisdictionSection|DE": "correct"},
                                 test_mode=test_mode)
    return ctx, corr


def test_golden_correction_structure_test_mode():
    ctx, corr = _golden_correction(test_mode=True)
    root = ET.fromstring(corr)
    assert root.find(N + "MessageSpec").find(N + "MessageTypeIndic").text == "GIR102"
    # fresh MessageRefId
    assert root.find(N + "MessageSpec").find(N + "MessageRefId").text != ctx["message_ref_id"]

    body = root.find(N + "GLOBEBody")
    tags = [G["_local_tag"](el.tag) for el in body]
    assert tags == ["FilingInfo", "JurisdictionSection"]   # unchanged sections OMITTED

    fi_dti, fi_dr, fi_corr = _docspec(body.find(N + "FilingInfo"))
    assert fi_dti == "OECD10"                              # test-flavor Resend
    assert fi_dr == ctx["filing_doc_ref_id"]               # rule 60014: SAME DocRefId
    assert fi_corr is None                                 # rule 60012

    orig_de = next(s for s in ctx["sections"] if s["key"] == "JurisdictionSection|DE")
    dti, dr, corr_ref = _docspec(body.find(N + "JurisdictionSection"))
    assert dti == "OECD12"                                 # test-flavor correction
    assert corr_ref == orig_de["doc_ref_id"]               # chain to the original
    assert dr != orig_de["doc_ref_id"] and re.match(r"^CH2024-", dr)
    rec = body.find(N + "JurisdictionSection").find(N + "RecJurCode")
    assert rec.text == orig_de["rec_jur_code"]             # rule 98200


def test_golden_correction_prod_mode_uses_oecd0_oecd2():
    _ctx, corr = _golden_correction(test_mode=False)
    body = ET.fromstring(corr).find(N + "GLOBEBody")
    assert _docspec(body.find(N + "FilingInfo"))[0] == "OECD0"
    assert _docspec(body.find(N + "JurisdictionSection"))[0] == "OECD2"


def test_golden_correction_validates_green_both_layers():
    ctx, corr = _golden_correction()
    l1, l2 = G["validate_correction"](corr, ctx)
    assert [c for c in l1 if not c[1]] == []
    assert [c for c in l2 if not c[1]] == []


def test_merged_view_carries_the_corrected_value():
    computations, cfg = _cfg_from_excel(NORDSTREAM)
    original = G["build_xml"](computations, cfg, test_mode=True)
    ctx = G["parse_original_xml"](original.encode())
    comp2 = copy.deepcopy(computations)
    de = next(c for c in comp2 if c["jur_iso"] == "DE")
    de["adjusted_fanil"] += 1
    new_xml = G["build_xml"](comp2, cfg, test_mode=True)
    corr = G["apply_correction"](new_xml, ctx, {"JurisdictionSection|DE": "correct"}, True)
    merged = G["build_merged_view"](ctx, corr)
    mroot = ET.fromstring(merged)
    de_sec = next(el for el in mroot.find(N + "GLOBEBody")
                  if G["_local_tag"](el.tag) == "JurisdictionSection"
                  and G["_sec_jurisdiction"](el) == "DE")
    fanil = de_sec.find(".//" + N + "FANIL")
    assert fanil.text == str(de["adjusted_fanil"])
    # merged report is still fully valid on the existing structural checks
    assert [c for c in G["validate_xml"](merged) if not c[1]] == []


def test_storno_section_from_original():
    """Storno of a section: OECD13 (test), CorrDocRefId set, gone from the merged view."""
    computations, cfg = _cfg_from_excel(NORDSTREAM)
    original = G["build_xml"](computations, cfg, test_mode=True)
    ctx = G["parse_original_xml"](original.encode())
    new_xml = G["build_xml"](computations, cfg, test_mode=True)
    corr = G["apply_correction"](new_xml, ctx, {"Summary|SE": "delete"}, test_mode=True)
    body = ET.fromstring(corr).find(N + "GLOBEBody")
    tags = [G["_local_tag"](el.tag) for el in body]
    assert tags == ["FilingInfo", "Summary"]
    dti, _dr, corr_ref = _docspec(body.find(N + "Summary"))
    orig_se = next(s for s in ctx["sections"] if s["key"] == "Summary|SE")
    assert dti == "OECD13" and corr_ref == orig_se["doc_ref_id"]
    merged = G["build_merged_view"](ctx, corr)
    assert "Summary|SE" not in _sections(merged)


def test_correction_on_firma1_single_computation():
    """Firma-1 shape (PureFert template): correct the CH computation only."""
    computations, cfg = _cfg_from_excel(PUREFERT)
    original = G["build_xml"](computations, cfg, test_mode=True)
    ctx = G["parse_original_xml"](original.encode())
    comp2 = copy.deepcopy(computations)
    comp2[0]["adjusted_fanil"] += 1
    new_xml = G["build_xml"](comp2, cfg, test_mode=True)
    key = f"JurisdictionSection|{comp2[0]['jur_iso'] or 'CH'}"
    rows = G["diff_correction_sections"](ctx, new_xml)
    assert [r["key"] for r in rows if r["status"] == "changed"] == [key]
    corr = G["apply_correction"](new_xml, ctx, {key: "correct"}, test_mode=True)
    l1, l2 = G["validate_correction"](corr, ctx)
    assert [c for c in l1 if not c[1]] == []
    assert [c for c in l2 if not c[1]] == []


# ─── negative tests: every Layer-1 rule must FAIL when violated ───────────────

def _tamper(corr_xml: str, fn) -> str:
    root = ET.fromstring(corr_xml)
    fn(root)
    return ET.tostring(root, encoding="unicode")


def _failing_labels(ctx, xml_str):
    return {label for label, ok, _ in G["validate_correction_xml"](xml_str, ctx) if not ok}


def test_negative_wrong_message_type():
    ctx, corr = _golden_correction()
    bad = _tamper(corr, lambda r: setattr(
        r.find(N + "MessageSpec").find(N + "MessageTypeIndic"), "text", "GIR101"))
    assert any("GIR102" in l for l in _failing_labels(ctx, bad))


def test_negative_missing_corrdocrefid():
    ctx, corr = _golden_correction()

    def rm(root):
        ds = root.find(N + "GLOBEBody").find(N + "JurisdictionSection").find(N + "DocSpec")
        ds.remove(ds.find(S + "CorrDocRefId"))
    assert any("60015" in l for l in _failing_labels(ctx, _tamper(corr, rm)))


def test_negative_reused_docrefid():
    ctx, corr = _golden_correction()

    def reuse(root):
        ds = root.find(N + "GLOBEBody").find(N + "JurisdictionSection").find(N + "DocSpec")
        ds.find(S + "DocRefId").text = ds.find(S + "CorrDocRefId").text
    assert any("60007" in l for l in _failing_labels(ctx, _tamper(corr, reuse)))


def test_negative_changed_recjurcode():
    ctx, corr = _golden_correction()

    def flip(root):
        root.find(N + "GLOBEBody").find(N + "JurisdictionSection").find(N + "RecJurCode").text = "DE"
    assert any("98200" in l for l in _failing_labels(ctx, _tamper(corr, flip)))


def test_negative_filinginfo_wrong_docrefid():
    ctx, corr = _golden_correction()

    def swap(root):
        ds = root.find(N + "GLOBEBody").find(N + "FilingInfo").find(N + "DocSpec")
        ds.find(S + "DocRefId").text = "CH2024-not-the-original"
    assert any("60014" in l for l in _failing_labels(ctx, _tamper(corr, swap)))


def test_negative_new_section_mixed_in():
    ctx, corr = _golden_correction()

    def mix(root):
        ds = root.find(N + "GLOBEBody").find(N + "JurisdictionSection").find(N + "DocSpec")
        ds.find(S + "DocTypeIndic").text = "OECD11"
    assert any("60004" in l for l in _failing_labels(ctx, _tamper(corr, mix)))


def test_negative_corrdocref_wrong_type():
    ctx, corr = _golden_correction()
    gen_ref = next(s["doc_ref_id"] for s in ctx["sections"] if s["type"] == "GeneralSection")

    def wrong(root):
        ds = root.find(N + "GLOBEBody").find(N + "JurisdictionSection").find(N + "DocSpec")
        ds.find(S + "CorrDocRefId").text = gen_ref
    assert any("60005" in l for l in _failing_labels(ctx, _tamper(corr, wrong)))


def test_negative_test_indic_mixed_with_prod():
    ctx, corr = _golden_correction(test_mode=True)

    def mix(root):
        ds = root.find(N + "GLOBEBody").find(N + "FilingInfo").find(N + "DocSpec")
        ds.find(S + "DocTypeIndic").text = "OECD0"     # prod resend in a test file
    assert any("50009" in l for l in _failing_labels(ctx, _tamper(corr, mix)))


# ─── regression guards ────────────────────────────────────────────────────────

def test_neumeldung_pipeline_untouched_all_checks_green():
    """Parity guard: the Neumeldung path (both templates) still passes every
    existing structural check — corrections are pure post-processing."""
    for tpl in (PUREFERT, NORDSTREAM):
        computations, cfg = _cfg_from_excel(tpl)
        xml = G["build_xml"](computations, cfg, test_mode=True)
        failed = [c for c in G["validate_xml"](xml) if not c[1]]
        assert failed == [], f"{tpl.name}: {failed}"


def test_correction_file_is_xsd_valid():
    lxml_etree = pytest.importorskip("lxml.etree")
    schema = lxml_etree.XMLSchema(lxml_etree.parse(str(XSD)))
    _ctx, corr = _golden_correction()
    doc = lxml_etree.fromstring(corr.encode())
    assert schema.validate(doc), schema.error_log
