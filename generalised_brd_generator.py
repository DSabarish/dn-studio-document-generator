#!/usr/bin/env python3
"""
generalised_brd_generator.py (v6 — FR bare-list fix)

Fixes vs v5:
  Issue FR — Functional Requirements: sub-sections (6.1, 6.2 …) are BARE LISTS
             at the top level of the FR dict, NOT wrapped in {value: [...]}.
             Each item is a dict with keys:
               "Requirement ID", "Description" (or "value"), "Priority", "trace"
             Parser now handles this shape correctly.
             Uses "Description" as the requirement text field (not "Requirement").
             Preserves original "Requirement ID" from JSON (FR-001, FR-002 …).

  Issue NFR — Section keys may include non-"6." prefixed extras (INSTRUCTION,
              Priority Rating Scale). Parser now filters strictly on keys that
              start with "6." when iterating FR sections.

why_null display messages (must match brd_llm.py):
  not_mentioned           → "Not discussed in the meeting — to be captured in a future session."
  deferred_by_participants → "Raised in the meeting but explicitly deferred — participants agreed to decide later."
  not_yet_agreed          → "Discussed in the meeting but not resolved — consensus was not reached."
  schema_excluded         → "Not applicable to this project — no evidence found in the transcript for this category."
  insufficient_detail     → "Mentioned in the meeting but without sufficient detail to document a concrete requirement."
  not_applicable          → "Not applicable to this project type."
"""

import json
import sys
import subprocess
import argparse
from pathlib import Path


# ---------------------------------------------------------------------------
# why_null display messages
# ---------------------------------------------------------------------------
WHY_NULL_DISPLAY = {
    "not_mentioned":             "Not discussed in the meeting — to be captured in a future session.",
    "deferred_by_participants":  "Raised in the meeting but explicitly deferred — participants agreed to decide later.",
    "not_yet_agreed":            "Discussed in the meeting but not resolved — consensus was not reached.",
    "schema_excluded":           "Not applicable to this project — no evidence found in the transcript for this category.",
    "insufficient_detail":       "Mentioned in the meeting but without sufficient detail to document a concrete requirement.",
    "not_applicable":            "Not applicable to this project type.",
}
WHY_NULL_FALLBACK = "Not recorded — to be completed before document review."


def _why_null_message(obj) -> str:
    code = obj.get("why_null") if isinstance(obj, dict) else None
    return WHY_NULL_DISPLAY.get(code, WHY_NULL_FALLBACK)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def v(obj, *keys, default=None):
    """Traverse nested dict and return innermost .value, or None."""
    cur = obj
    for k in keys:
        if not isinstance(cur, dict):
            return default
        cur = cur.get(k, {})
    if isinstance(cur, dict):
        return cur.get("value", default)
    return cur if cur is not None else default


def v_obj(obj, *keys):
    """Return the raw dict at the key path (for why_null access)."""
    cur = obj
    for k in keys:
        if not isinstance(cur, dict):
            return {}
        cur = cur.get(k, {})
    return cur if isinstance(cur, dict) else {}


def vlist(obj, *keys, default=None):
    result = v(obj, *keys, default=default if default is not None else [])
    if isinstance(result, list):
        return result
    if result is None:
        return []
    return [str(result)]


def vdict(obj, *keys, default=None):
    result = v(obj, *keys, default=default if default is not None else {})
    if isinstance(result, dict):
        return result
    return {}


def safe_str(val, default="—"):
    if val is None or val == "" or (isinstance(val, list) and len(val) == 0):
        return default
    if isinstance(val, list):
        return "; ".join(str(x) for x in val)
    return str(val)


def js_str(s):
    if s is None:
        s = ""
    s = str(s)
    s = s.replace("\\", "\\\\")
    s = s.replace('"', '\\"')
    s = s.replace("\n", " ")
    s = s.replace("\r", "")
    return s


def _get_section(data: dict, key: str, fallback_parent: str = None) -> dict:
    if key in data:
        return data[key]
    if fallback_parent and fallback_parent in data:
        parent = data[fallback_parent]
        if isinstance(parent, dict) and key in parent:
            return parent[key]
    return {}


# ---------------------------------------------------------------------------
# Parse JSON
# ---------------------------------------------------------------------------
def parse_brd(data: dict) -> dict:
    H  = data.get("Header", {})
    P  = data.get("0. Project Details", {})
    DR = _get_section(data, "0. Document Revisions", "0. Project Details")
    SA = _get_section(data, "0. Stakeholder Approvals", "0. Project Details")
    I  = data.get("1. Introduction", {})
    S  = data.get("2. Project Scope", {})
    SP = data.get("3. System Perspective", {})
    BP = data.get("4. Business Process Overview", {})
    KP = data.get("5. KPI & Success Metrics", {})
    FR = data.get("6. Functional Requirements", {})
    NF = data.get("7. Non-Functional Requirements", {})
    DG = data.get("8. Data Governance & Privacy", {})
    TA = data.get("9. Technology Stack & Architecture", {})
    AP = data.get("10. Appendices", {})

    ctx = {}

    # ── Header ──────────────────────────────────────────────────────────────
    ctx["project_name"]   = safe_str(v(H, "Project Name"))
    ctx["version"]        = safe_str(v(H, "Version"))
    ctx["date"]           = safe_str(v(H, "Date"))
    ctx["business_unit"]  = safe_str(v(H, "Business Unit"))
    ctx["project_type"]   = safe_str(v(H, "Project Type"))
    ctx["deployment"]     = safe_str(v(H, "Deployment Model"))
    ctx["data_store"]     = safe_str(v(H, "Primary Data Store"))
    ctx["compute"]        = safe_str(v(H, "Compute Engine"))
    ctx["infrastructure"] = safe_str(v(H, "Infrastructure"))
    ctx["doc_status"]     = safe_str(v(H, "Document Status"))

    # ── Project Details ──────────────────────────────────────────────────────
    ctx["overview"]             = v(P, "Project Overview")
    ctx["overview_why"]         = _why_null_message(v_obj(P, "Project Overview"))
    ctx["business_need"]        = v(P, "Business Need")
    ctx["business_need_why"]    = _why_null_message(v_obj(P, "Business Need"))
    ctx["success_criteria"]     = vlist(P, "Success Criteria")
    ctx["success_criteria_why"] = _why_null_message(v_obj(P, "Success Criteria"))

    # ── Document Revisions ───────────────────────────────────────────────────
    raw_revisions = DR.get("value") or DR.get("items") or []
    ctx["revisions"] = []
    for r in raw_revisions:
        if isinstance(r, dict):
            ctx["revisions"].append({
                "version": safe_str(v(r, "Version")),
                "date":    safe_str(v(r, "Date")),
                "author":  safe_str(v(r, "Author")),
                "desc":    safe_str(v(r, "Description")),
            })

    # ── Stakeholder Approvals ────────────────────────────────────────────────
    raw_approvals = SA.get("value") or SA.get("items") or []
    ctx["approvals"] = []
    for a in raw_approvals:
        if isinstance(a, dict):
            ctx["approvals"].append({
                "name":   safe_str(v(a, "Name")),
                "role":   safe_str(v(a, "Role")),
                "status": safe_str(v(a, "Status")),
            })

    # ── Introduction ─────────────────────────────────────────────────────────
    ctx["project_summary"]     = v(I, "1.1 Project Summary")
    ctx["project_summary_why"] = _why_null_message(v_obj(I, "1.1 Project Summary"))

    obj_block = I.get("1.2 Objectives", {})
    obj_items = obj_block.get("value") or obj_block.get("items") or {}
    ctx["obj_primary"]       = vlist(obj_items, "Primary")
    ctx["obj_primary_why"]   = _why_null_message(v_obj(obj_items, "Primary"))
    ctx["obj_secondary"]     = vlist(obj_items, "Secondary")
    ctx["obj_secondary_why"] = _why_null_message(v_obj(obj_items, "Secondary"))

    ctx["background"]           = v(I, "1.3 Background & Business Context")
    ctx["background_why"]       = _why_null_message(v_obj(I, "1.3 Background & Business Context"))
    ctx["business_drivers"]     = vlist(I, "1.4 Business Drivers")
    ctx["business_drivers_why"] = _why_null_message(v_obj(I, "1.4 Business Drivers"))

    # ── Scope ────────────────────────────────────────────────────────────────
    ctx["in_scope"]      = vlist(S, "2.1 In-Scope Functionality")
    ctx["in_scope_why"]  = _why_null_message(v_obj(S, "2.1 In-Scope Functionality"))
    ctx["out_scope"]     = vlist(S, "2.2 Out-of-Scope Functionality")
    ctx["out_scope_why"] = _why_null_message(v_obj(S, "2.2 Out-of-Scope Functionality"))

    raw_phases = S.get("2.3 Phasing Plan", {}).get("value") or S.get("2.3 Phasing Plan", {}).get("items") or {}
    ctx["phases"] = []
    for phase_key in sorted(raw_phases.keys()):
        phase_obj = raw_phases[phase_key]
        if not isinstance(phase_obj, dict):
            ctx["phases"].append((phase_key, safe_str(phase_obj), bool(phase_obj)))
            continue
        phase_val = phase_obj.get("value")
        phase_why = _why_null_message(phase_obj)
        has_value = phase_val is not None and phase_val != ""
        display = safe_str(phase_val) if has_value else phase_why
        ctx["phases"].append((phase_key, display, has_value))

    # ── System Perspective ───────────────────────────────────────────────────
    ctx["assumptions"]     = vlist(SP, "3.1 Assumptions")
    ctx["assumptions_why"] = _why_null_message(v_obj(SP, "3.1 Assumptions"))
    ctx["constraints"]     = vlist(SP, "3.2 Constraints")
    ctx["constraints_why"] = _why_null_message(v_obj(SP, "3.2 Constraints"))

    risks_obj = SP.get("3.3 Risks", {})
    ctx["risks_why"] = _why_null_message(risks_obj)
    ctx["risks"] = []
    raw_risks_val = risks_obj.get("value") or risks_obj.get("items") or []
    if isinstance(raw_risks_val, list):
        for r in raw_risks_val:
            if isinstance(r, dict):
                risk_text = r.get("Risk") or r.get("risk") or ""
                mit_text  = r.get("Mitigation") or r.get("mitigation") or ""
                if isinstance(risk_text, dict): risk_text = risk_text.get("value", "")
                if isinstance(mit_text,  dict): mit_text  = mit_text.get("value", "")
                ctx["risks"].append([safe_str(risk_text), safe_str(mit_text)])
    elif isinstance(raw_risks_val, dict):
        risk_text = raw_risks_val.get("Risk") or raw_risks_val.get("risk") or {}
        mit_text  = raw_risks_val.get("Mitigation") or raw_risks_val.get("mitigation") or {}
        if isinstance(risk_text, dict): risk_text = risk_text.get("value", "")
        if isinstance(mit_text,  dict): mit_text  = mit_text.get("value", "")
        if risk_text or mit_text:
            ctx["risks"].append([safe_str(risk_text), safe_str(mit_text)])

    # ── Business Process ─────────────────────────────────────────────────────
    ctx["as_is"]      = vlist(BP, "4.1 Current Process (As-Is)")
    ctx["as_is_why"]  = _why_null_message(v_obj(BP, "4.1 Current Process (As-Is)"))
    ctx["to_be"]      = vlist(BP, "4.2 Proposed Process (To-Be)")
    ctx["to_be_why"]  = _why_null_message(v_obj(BP, "4.2 Proposed Process (To-Be)"))

    # ── KPIs ─────────────────────────────────────────────────────────────────
    kp_val = KP.get("value") or KP.get("items") or []
    ctx["kpis"] = []
    ctx["kpis_why"] = _why_null_message(KP)
    if isinstance(kp_val, list):
        for k in kp_val:
            if isinstance(k, dict):
                metric = k.get("Metric") or k.get("metric") or {}
                target = k.get("Target") or k.get("target") or {}
                if isinstance(metric, dict): metric = metric.get("value", "")
                if isinstance(target, dict): target = target.get("value") or ""
                ctx["kpis"].append([safe_str(metric), safe_str(target) if target else "—"])

    # ── Functional Requirements (v6 fix) ─────────────────────────────────────
    # The JSON shape for sub-sections is a BARE LIST at the top level of FR,
    # NOT wrapped in {value: [...]}. Each list item is a dict:
    #   { "Requirement ID": "FR-001", "Description": "...", "Priority": "P1",
    #     "value": "...", "trace": {...} }
    # We filter only keys that start with "6." to skip INSTRUCTION / Priority
    # Rating Scale entries.
    ctx["fr_sections"] = []
    for sec_key in sorted(fk for fk in FR.keys() if fk.startswith("6.")):
        sec_obj = FR[sec_key]

        # ── determine raw_reqs list and why message ──────────────────────────
        if isinstance(sec_obj, list):
            # PRIMARY SHAPE in this JSON: bare list
            raw_reqs = sec_obj
            sec_why  = WHY_NULL_FALLBACK
        elif isinstance(sec_obj, dict):
            # FALLBACK SHAPE: {value: [...], why_null: "..."}
            inner = sec_obj.get("value") or sec_obj.get("items") or []
            raw_reqs = inner if isinstance(inner, list) else []
            sec_why  = _why_null_message(sec_obj)
        else:
            raw_reqs = []
            sec_why  = WHY_NULL_FALLBACK

        reqs = []
        if raw_reqs:
            try:
                sub    = int(sec_key.split(".")[1])
                prefix = f"F{sub:02d}"
            except (IndexError, ValueError):
                prefix = "REQ"

            for idx, req in enumerate(raw_reqs, 1):
                if not isinstance(req, dict):
                    # plain string fallback
                    req_str  = str(req)
                    priority = "P1"
                    for tag in ["P1", "P2", "P3"]:
                        if f"({tag}" in req_str or f"[{tag}]" in req_str:
                            priority = tag
                            break
                    clean = req_str
                    for tag in ["P1", "P2", "P3"]:
                        clean = clean.replace(f"[{tag}]", "").replace(f"({tag})", "").strip()
                    reqs.append([f"{prefix}-{idx:03d}", clean.rstrip("."), priority])
                    continue

                # ── Requirement ID ───────────────────────────────────────────
                req_id = (
                    req.get("Requirement ID") or
                    req.get("requirement_id") or
                    f"{prefix}-{idx:03d}"
                )

                # ── Requirement text ─────────────────────────────────────────
                # Prefer "Description" (this JSON's field), then fallbacks
                req_text = (
                    req.get("Description") or
                    req.get("Requirement") or
                    req.get("requirement") or
                    req.get("value") or
                    ""
                )
                if isinstance(req_text, dict):
                    req_text = req_text.get("value", "")

                # ── Priority ─────────────────────────────────────────────────
                priority = req.get("Priority") or req.get("priority") or "P1"
                if isinstance(priority, dict):
                    priority = priority.get("value", "P1")
                priority = str(priority).strip("[]() ")
                if priority not in ("P1", "P2", "P3"):
                    priority = "P1"

                reqs.append([str(req_id), str(req_text).rstrip("."), priority])

        ctx["fr_sections"].append({
            "title": sec_key,
            "key":   sec_key,
            "reqs":  reqs,
            "why":   sec_why,
        })

    # ── Non-Functional Requirements ──────────────────────────────────────────
    # NFR shape is {value: [...], why_null: "..."} — vlist handles this correctly.
    for ctx_key, nfr_section in [
        ("nfr_perf",       "7.1 Performance & Scalability"),
        ("nfr_avail",      "7.2 Availability & Reliability"),
        ("nfr_usab",       "7.3 Usability & Accessibility"),
        ("nfr_sec",        "7.4 Security & Access Control"),
        ("nfr_compliance", "7.5 Compliance & Regulatory"),
    ]:
        ctx[ctx_key]          = vlist(NF, nfr_section)
        ctx[ctx_key + "_why"] = _why_null_message(v_obj(NF, nfr_section))

    # ── Data Governance ──────────────────────────────────────────────────────
    raw_dc = vdict(DG, "8.1 Data Classification")
    ctx["data_classification"]     = [[k, safe_str(vv)] for k, vv in raw_dc.items()] if raw_dc else []
    ctx["data_classification_why"] = _why_null_message(v_obj(DG, "8.1 Data Classification"))
    ctx["privacy_checklist"]       = vlist(DG, "8.2 Data Privacy Checklist")
    ctx["privacy_checklist_why"]   = _why_null_message(v_obj(DG, "8.2 Data Privacy Checklist"))

    # ── Tech Stack ───────────────────────────────────────────────────────────
    raw_ts = vdict(TA, "9.1 Proposed Technology Stack")
    ctx["tech_stack"]     = [[k, safe_str(vv)] for k, vv in raw_ts.items()] if raw_ts else []
    ctx["tech_stack_why"] = _why_null_message(v_obj(TA, "9.1 Proposed Technology Stack"))

    # ── Appendices ───────────────────────────────────────────────────────────
    raw_glossary = vdict(AP, "10.1 Glossary of Terms")
    ctx["glossary"]         = sorted([[k, safe_str(vv)] for k, vv in raw_glossary.items()]) if raw_glossary else []
    ctx["glossary_why"]     = _why_null_message(v_obj(AP, "10.1 Glossary of Terms"))
    ctx["acronyms"]         = vlist(AP, "10.2 List of Acronyms")
    ctx["acronyms_why"]     = _why_null_message(v_obj(AP, "10.2 List of Acronyms"))
    ctx["related_docs"]     = vlist(AP, "10.3 Related Documents")
    ctx["related_docs_why"] = _why_null_message(v_obj(AP, "10.3 Related Documents"))

    raw_signoff = AP.get("10.4 Document Sign-Off", {}).get("value") or []
    ctx["signoff"] = []
    for s in raw_signoff:
        if isinstance(s, dict):
            ctx["signoff"].append([
                safe_str(s.get("Name")),
                safe_str(s.get("Role")),
                safe_str(s.get("Date") or "________________"),
                safe_str(s.get("Signature") or "________________"),
            ])

    return ctx


# ---------------------------------------------------------------------------
# Build Node.js generator script
# ---------------------------------------------------------------------------
def build_js(ctx: dict) -> str:
    def jsa(lst):
        return "[" + ", ".join(f'"{js_str(x)}"' for x in lst) + "]"

    def jsa2(lst):
        return "[" + ", ".join(f'["{js_str(r[0])}", "{js_str(r[1])}"]' for r in lst) + "]"

    def jsa3(lst):
        return "[" + ", ".join(f'["{js_str(r[0])}", "{js_str(r[1])}", "{js_str(r[2])}"]' for r in lst) + "]"

    def jsa4(lst):
        return "[" + ", ".join(f'["{js_str(r[0])}", "{js_str(r[1])}", "{js_str(r[2])}", "{js_str(r[3])}"]' for r in lst) + "]"

    fr_js_parts = []
    for sec in ctx["fr_sections"]:
        fr_js_parts.append(
            f'{{ title: "{js_str(sec["key"])}", reqs: {jsa3(sec["reqs"])}, why: "{js_str(sec["why"])}" }}'
        )
    fr_sections_js = "[" + ",\n".join(fr_js_parts) + "]"

    phases_js_parts = []
    for (label, display, has_value) in ctx["phases"]:
        phases_js_parts.append(
            f'{{ label: "{js_str(label)}", desc: "{js_str(display)}", hasValue: {str(has_value).lower()} }}'
        )
    phases_js = "[" + ", ".join(phases_js_parts) + "]"

    rev_items = [
        f'["{js_str(r["version"])}", "{js_str(r["date"])}", "{js_str(r["author"])}", "{js_str(r["desc"])}"]'
        for r in ctx["revisions"]
    ]
    approvals_items = [
        f'["{js_str(a["name"])}", "{js_str(a["role"])}", "{js_str(a["status"])}"]'
        for a in ctx["approvals"]
    ]

    script = f"""
const {{ Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat, TableOfContents,
  TabStopType, TabStopPosition }} = require('docx');
const fs = require('fs');

const D = {{
  projectName: "{js_str(ctx['project_name'])}",
  version: "{js_str(ctx['version'])}",
  date: "{js_str(ctx['date'])}",
  businessUnit: "{js_str(ctx['business_unit'])}",
  projectType: "{js_str(ctx['project_type'])}",
  deployment: "{js_str(ctx['deployment'])}",
  dataStore: "{js_str(ctx['data_store'])}",
  compute: "{js_str(ctx['compute'])}",
  infrastructure: "{js_str(ctx['infrastructure'])}",
  docStatus: "{js_str(ctx['doc_status'])}",

  overview: {json.dumps(ctx['overview'])},
  overviewWhy: "{js_str(ctx['overview_why'])}",
  businessNeed: {json.dumps(ctx['business_need'])},
  businessNeedWhy: "{js_str(ctx['business_need_why'])}",
  successCriteria: {jsa(ctx['success_criteria'])},
  successCriteriaWhy: "{js_str(ctx['success_criteria_why'])}",

  revisions: [{", ".join(rev_items)}],
  approvals: [{", ".join(approvals_items)}],

  projectSummary: {json.dumps(ctx['project_summary'])},
  projectSummaryWhy: "{js_str(ctx['project_summary_why'])}",
  objPrimary: {jsa(ctx['obj_primary'])},
  objPrimaryWhy: "{js_str(ctx['obj_primary_why'])}",
  objSecondary: {jsa(ctx['obj_secondary'])},
  objSecondaryWhy: "{js_str(ctx['obj_secondary_why'])}",
  background: {json.dumps(ctx['background'])},
  backgroundWhy: "{js_str(ctx['background_why'])}",
  businessDrivers: {jsa(ctx['business_drivers'])},
  businessDriversWhy: "{js_str(ctx['business_drivers_why'])}",

  inScope: {jsa(ctx['in_scope'])},
  inScopeWhy: "{js_str(ctx['in_scope_why'])}",
  outScope: {jsa(ctx['out_scope'])},
  outScopeWhy: "{js_str(ctx['out_scope_why'])}",
  phases: {phases_js},

  assumptions: {jsa(ctx['assumptions'])},
  assumptionsWhy: "{js_str(ctx['assumptions_why'])}",
  constraints: {jsa(ctx['constraints'])},
  constraintsWhy: "{js_str(ctx['constraints_why'])}",
  risks: {jsa2(ctx['risks'])},
  risksWhy: "{js_str(ctx['risks_why'])}",

  asIs: {jsa(ctx['as_is'])},
  asIsWhy: "{js_str(ctx['as_is_why'])}",
  toBe: {jsa(ctx['to_be'])},
  toBeWhy: "{js_str(ctx['to_be_why'])}",

  kpis: {jsa2(ctx['kpis'])},
  kpisWhy: "{js_str(ctx['kpis_why'])}",

  frSections: {fr_sections_js},

  nfrPerf: {jsa(ctx['nfr_perf'])},
  nfrPerfWhy: "{js_str(ctx['nfr_perf_why'])}",
  nfrAvail: {jsa(ctx['nfr_avail'])},
  nfrAvailWhy: "{js_str(ctx['nfr_avail_why'])}",
  nfrUsab: {jsa(ctx['nfr_usab'])},
  nfrUsabWhy: "{js_str(ctx['nfr_usab_why'])}",
  nfrSec: {jsa(ctx['nfr_sec'])},
  nfrSecWhy: "{js_str(ctx['nfr_sec_why'])}",
  nfrCompliance: {jsa(ctx['nfr_compliance'])},
  nfrComplianceWhy: "{js_str(ctx['nfr_compliance_why'])}",

  dataClassification: {jsa2(ctx['data_classification'])},
  dataClassificationWhy: "{js_str(ctx['data_classification_why'])}",
  privacyChecklist: {jsa(ctx['privacy_checklist'])},
  privacyChecklistWhy: "{js_str(ctx['privacy_checklist_why'])}",

  techStack: {jsa2(ctx['tech_stack'])},
  techStackWhy: "{js_str(ctx['tech_stack_why'])}",

  glossary: {jsa2(ctx['glossary'])},
  glossaryWhy: "{js_str(ctx['glossary_why'])}",
  acronyms: {jsa(ctx['acronyms'])},
  acronymsWhy: "{js_str(ctx['acronyms_why'])}",
  relatedDocs: {jsa(ctx['related_docs'])},
  relatedDocsWhy: "{js_str(ctx['related_docs_why'])}",
  signoff: {jsa4(ctx['signoff'])},
}};

const C = {{
  primary: "003366", accent: "0070C0", lightBlue: "D6E4F0",
  dark: "404040", mid: "666666", border: "CCCCCC",
  white: "FFFFFF", placeholder: "DDDDDD", note: "888888",
}};
const PW = 9360;

const bd = (c=C.border) => ({{ style: BorderStyle.SINGLE, size: 1, color: c }});
const bds = (c=C.border) => ({{ top:bd(c), bottom:bd(c), left:bd(c), right:bd(c) }});

function sp(n=100) {{
  return new Paragraph({{ spacing:{{before:n, after:0}}, children:[new TextRun("")] }});
}}
function divider(color=C.accent) {{
  return new Paragraph({{
    border: {{ bottom: {{ style:BorderStyle.SINGLE, size:12, color, space:1 }} }},
    spacing: {{ before:0, after:200 }}, children: [new TextRun("")]
  }});
}}
function h1(text, pageBreak=true) {{
  return new Paragraph({{
    heading: HeadingLevel.HEADING_1, pageBreakBefore: pageBreak,
    children: [new TextRun({{ text, bold:true, color:C.primary }})],
    spacing: {{ before:200, after:120 }}
  }});
}}
function h2(text) {{
  return new Paragraph({{
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({{ text, bold:true, color:C.accent }})],
    spacing: {{ before:280, after:100 }}
  }});
}}
function h3(text) {{
  return new Paragraph({{
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({{ text, bold:true, color:C.dark, size:22 }})],
    spacing: {{ before:200, after:80 }}
  }});
}}
function body(text, {{bold=false, italic=false, color=C.dark, size=22}}={{}}) {{
  return new Paragraph({{
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({{ text, bold, italic, color, size, font:"Arial" }})],
    spacing: {{ before:60, after:60 }}
  }});
}}
function nullNote(whyMsg) {{
  return new Paragraph({{
    alignment: AlignmentType.LEFT,
    children: [
      new TextRun({{ text: "! ", color:C.note, size:20, font:"Arial" }}),
      new TextRun({{ text: whyMsg, italics:true, color:C.note, size:20, font:"Arial" }})
    ],
    spacing: {{ before:60, after:60 }}
  }});
}}
function bullet(text, ref) {{
  return new Paragraph({{
    numbering: {{ reference: ref || "b0", level:0 }},
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({{ text, size:22, color:C.dark, font:"Arial" }})],
    spacing: {{ before:40, after:40 }}
  }});
}}
function tbl(headers, rows, colWidths) {{
  const total = colWidths.reduce((a,b)=>a+b, 0);
  const hrow = new TableRow({{
    children: headers.map((h,i) => new TableCell({{
      borders: bds(), width: {{ size:colWidths[i], type:WidthType.DXA }},
      shading: {{ fill:C.primary, type:ShadingType.CLEAR }},
      margins: {{ top:100, bottom:100, left:150, right:150 }},
      children:[new Paragraph({{ children:[new TextRun({{text:h, bold:true, size:20, color:C.white, font:"Arial"}})] }})]
    }}))
  }});
  const drows = rows.map(row => new TableRow({{
    children: row.map((cell,i) => new TableCell({{
      borders: bds(), width: {{ size:colWidths[i], type:WidthType.DXA }},
      margins: {{ top:80, bottom:80, left:150, right:150 }},
      children:[new Paragraph({{ alignment:AlignmentType.JUSTIFIED,
        children:[new TextRun({{text:String(cell||"—"), size:20, color:C.dark, font:"Arial"}})] }})]
    }}))
  }}));
  return new Table({{ width:{{size:total, type:WidthType.DXA}}, columnWidths:colWidths, rows:[hrow,...drows] }});
}}
function kvTable(data) {{
  const rows = data.map(([k,val]) => new TableRow({{
    children:[
      new TableCell({{ borders:bds(), width:{{size:2400, type:WidthType.DXA}},
        shading:{{fill:C.lightBlue, type:ShadingType.CLEAR}},
        margins:{{top:100, bottom:100, left:150, right:150}},
        children:[new Paragraph({{children:[new TextRun({{text:k, bold:true, size:20, color:C.primary, font:"Arial"}})]}})]
      }}),
      new TableCell({{ borders:bds(), width:{{size:6960, type:WidthType.DXA}},
        margins:{{top:100, bottom:100, left:150, right:150}},
        children:[new Paragraph({{children:[new TextRun({{text:val, size:20, color:C.dark, font:"Arial"}})]}})]
      }}),
    ]
  }}));
  return new Table({{ width:{{size:PW, type:WidthType.DXA}}, columnWidths:[2400,6960], rows }});
}}
function archPlaceholder() {{
  const cell = new TableCell({{
    borders: bds(C.placeholder), width: {{ size:PW, type:WidthType.DXA }},
    shading: {{ fill:C.placeholder, type:ShadingType.CLEAR }},
    margins: {{ top:800, bottom:800, left:300, right:300 }},
    children:[
      new Paragraph({{ alignment:AlignmentType.CENTER, spacing:{{before:0,after:160}},
        children:[new TextRun({{text:"[ Attach Architecture Data Flow Diagram ]", bold:true, size:24, color:"888888", font:"Arial", italics:true}})]
      }}),
      new Paragraph({{ alignment:AlignmentType.CENTER,
        children:[new TextRun({{text:"Replace this placeholder with the finalised diagram before distributing.", size:18, color:"999999", font:"Arial", italics:true}})]
      }})
    ]
  }});
  return new Table({{ width:{{size:PW, type:WidthType.DXA}}, columnWidths:[PW], rows:[new TableRow({{children:[cell]}})] }});
}}

const bulletRefs = Array.from({{length:20}},(_,i)=>"b"+i);
const numberingConfig = bulletRefs.map(ref => ({{
  reference: ref,
  levels:[{{ level:0, format:LevelFormat.BULLET, text:"\\u2022", alignment:AlignmentType.LEFT,
    style:{{ paragraph:{{ indent:{{ left:540, hanging:300 }} }} }}
  }}]
}}));

const ch = [];

// COVER
ch.push(new Paragraph({{ spacing:{{before:1800,after:0}}, children:[new TextRun("")] }}));
ch.push(new Paragraph({{ alignment:AlignmentType.CENTER,
  children:[new TextRun({{text:"BUSINESS REQUIREMENTS DOCUMENT", bold:true, size:52, color:C.primary, font:"Arial"}})]
}}));
ch.push(sp(200));
ch.push(new Paragraph({{ alignment:AlignmentType.CENTER,
  children:[new TextRun({{text:D.projectName, bold:true, size:40, color:C.accent, font:"Arial"}})]
}}));
ch.push(sp(300));
ch.push(new Paragraph({{ alignment:AlignmentType.CENTER,
  border:{{bottom:{{style:BorderStyle.SINGLE,size:6,color:C.accent,space:1}}}},
  children:[new TextRun("")]
}}));
ch.push(sp(200));
[["Prepared for",D.businessUnit],["Version",D.version],["Date",D.date],["Document Status",D.docStatus]].forEach(([k,val])=>{{
  ch.push(new Paragraph({{ alignment:AlignmentType.CENTER, spacing:{{before:80,after:80}},
    children:[new TextRun({{text:k+": ",bold:true,size:22,color:C.mid,font:"Arial"}}),
              new TextRun({{text:val,size:22,color:C.dark,font:"Arial"}})]
  }}));
}});
ch.push(new Paragraph({{children:[new PageBreak()]}}));

// TOC
ch.push(h1("Table of Contents", false));
ch.push(divider());
ch.push(new TableOfContents("Table of Contents", {{hyperlink:true, headingStyleRange:"1-3"}}));
ch.push(new Paragraph({{children:[new PageBreak()]}}));

// DOCUMENT CONTROL
ch.push(h1("Document Control", false));
ch.push(divider());
ch.push(h3("Document Information"));
ch.push(sp(60));
ch.push(kvTable([
  ["Project Name",D.projectName],["Version",D.version],["Date",D.date],
  ["Business Unit",D.businessUnit],["Project Type",D.projectType],
  ["Deployment",D.deployment],["Data Store",D.dataStore],
  ["Compute Engine",D.compute],["Infrastructure",D.infrastructure],["Document Status",D.docStatus],
]));
ch.push(sp(200));
ch.push(h3("Document Revisions"));
ch.push(sp(60));
if (D.revisions.length > 0) {{
  ch.push(tbl(["Version","Date","Author","Description"], D.revisions, [900,1500,1800,5160]));
}} else {{
  ch.push(body("No revisions recorded.", {{italic:true, color:C.mid}}));
}}
ch.push(sp(200));
ch.push(h3("Stakeholder Approvals"));
ch.push(sp(60));
if (D.approvals.length > 0) {{
  ch.push(tbl(["Name","Role","Status"], D.approvals, [2400,4560,2400]));
}} else {{
  ch.push(body("No stakeholder approvals recorded.", {{italic:true, color:C.mid}}));
}}

// SECTION 1
ch.push(h1("1. Introduction"));
ch.push(divider());
ch.push(h2("1.1 Project Summary"));
D.projectSummary ? ch.push(body(D.projectSummary)) : ch.push(nullNote(D.projectSummaryWhy));
ch.push(sp(120));
ch.push(h2("1.2 Objectives"));
ch.push(body("Primary Objectives", {{bold:true, color:C.primary}}));
D.objPrimary.length > 0 ? D.objPrimary.forEach(x=>ch.push(bullet(x,"b0"))) : ch.push(nullNote(D.objPrimaryWhy));
ch.push(sp(80));
ch.push(body("Secondary Objectives", {{bold:true, color:C.accent}}));
D.objSecondary.length > 0 ? D.objSecondary.forEach(x=>ch.push(bullet(x,"b1"))) : ch.push(nullNote(D.objSecondaryWhy));
ch.push(sp(120));
ch.push(h2("1.3 Background & Business Context"));
D.background ? ch.push(body(D.background)) : ch.push(nullNote(D.backgroundWhy));
ch.push(sp(120));
ch.push(h2("1.4 Business Drivers"));
D.businessDrivers.length > 0 ? D.businessDrivers.forEach(x=>ch.push(bullet(x,"b2"))) : ch.push(nullNote(D.businessDriversWhy));

// SECTION 2
ch.push(h1("2. Project Scope"));
ch.push(divider());
ch.push(h2("2.1 In-Scope Functionality"));
D.inScope.length > 0 ? D.inScope.forEach(x=>ch.push(bullet(x,"b3"))) : ch.push(nullNote(D.inScopeWhy));
ch.push(sp(120));
ch.push(h2("2.2 Out-of-Scope Functionality"));
D.outScope.length > 0 ? D.outScope.forEach(x=>ch.push(bullet(x,"b4"))) : ch.push(nullNote(D.outScopeWhy));
ch.push(sp(120));
ch.push(h2("2.3 Phasing Plan"));
ch.push(sp(60));
if (D.phases.length > 0) {{
  const phHdr = new TableRow({{
    children: ["Phase","Scope Description"].map((h,i) => new TableCell({{
      borders:bds(), width:{{size:[1400,7960][i], type:WidthType.DXA}},
      shading:{{fill:C.primary, type:ShadingType.CLEAR}},
      margins:{{top:100,bottom:100,left:150,right:150}},
      children:[new Paragraph({{children:[new TextRun({{text:h,bold:true,size:20,color:C.white,font:"Arial"}})]}})]
    }}))
  }});
  const phRows = D.phases.map(p => new TableRow({{
    children: [
      new TableCell({{
        borders:bds(), width:{{size:1400,type:WidthType.DXA}},
        margins:{{top:80,bottom:80,left:150,right:150}},
        children:[new Paragraph({{children:[new TextRun({{text:p.label,size:20,color:C.dark,font:"Arial"}})]}})]
      }}),
      new TableCell({{
        borders:bds(), width:{{size:7960,type:WidthType.DXA}},
        margins:{{top:80,bottom:80,left:150,right:150}},
        children:[new Paragraph({{
          alignment: p.hasValue ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
          children:[
            p.hasValue
              ? new TextRun({{text:p.desc, size:20, color:C.dark, font:"Arial"}})
              : new TextRun({{text:"! "+p.desc, size:20, color:C.note, italics:true, font:"Arial"}})
          ]
        }})]
      }}),
    ]
  }}));
  ch.push(new Table({{width:{{size:PW,type:WidthType.DXA}},columnWidths:[1400,7960],rows:[phHdr,...phRows]}}));
}} else {{
  ch.push(nullNote("Not discussed in the meeting — to be captured in a future session."));
}}

// SECTION 3
ch.push(h1("3. System Perspective"));
ch.push(divider());
ch.push(h2("3.1 Assumptions"));
D.assumptions.length > 0 ? D.assumptions.forEach(x=>ch.push(bullet(x,"b5"))) : ch.push(nullNote(D.assumptionsWhy));
ch.push(sp(120));
ch.push(h2("3.2 Constraints"));
D.constraints.length > 0 ? D.constraints.forEach(x=>ch.push(bullet(x,"b6"))) : ch.push(nullNote(D.constraintsWhy));
ch.push(sp(120));
ch.push(h2("3.3 Risks & Mitigations"));
ch.push(sp(60));
D.risks.length > 0 ? ch.push(tbl(["Risk","Mitigation"],D.risks,[4680,4680])) : ch.push(nullNote(D.risksWhy));

// SECTION 4
ch.push(h1("4. Business Process Overview"));
ch.push(divider());
ch.push(h2("4.1 Current Process (As-Is)"));
D.asIs.length > 0 ? D.asIs.forEach(x=>ch.push(bullet(x,"b7"))) : ch.push(nullNote(D.asIsWhy));
ch.push(sp(120));
ch.push(h2("4.2 Proposed Process (To-Be)"));
D.toBe.length > 0 ? D.toBe.forEach(x=>ch.push(bullet(x,"b8"))) : ch.push(nullNote(D.toBeWhy));

// SECTION 5
ch.push(h1("5. KPI & Success Metrics"));
ch.push(divider());
ch.push(body("The following quantitative performance targets will be used to evaluate whether the system is operating successfully in production."));
ch.push(sp(80));
D.kpis.length > 0 ? ch.push(tbl(["Metric","Target"],D.kpis,[7200,2160])) : ch.push(nullNote(D.kpisWhy));

// SECTION 6
ch.push(h1("6. Functional Requirements"));
ch.push(divider());
ch.push(h3("Priority Rating Scale"));
ch.push(sp(60));
ch.push(tbl(["Priority","Definition"],[
  ["P1 — Must Have","Project is not complete without this. Blocking requirements — if not delivered, the project fails UAT."],
  ["P2 — Should Have","Strongly desired and planned, but project can launch without it if necessary. Deferral requires explicit stakeholder agreement."],
  ["P3 — Nice to Have","Valuable enhancement but not planned for current phase. Included for visibility and future planning only."],
],[2200,7160]));
ch.push(sp(120));
D.frSections.forEach(sec => {{
  ch.push(h2(sec.title));
  if (sec.reqs && sec.reqs.length > 0) {{
    ch.push(sp(60));
    ch.push(tbl(["ID","Requirement","Priority"],sec.reqs,[900,7460,1000]));
  }} else {{
    ch.push(nullNote(sec.why));
  }}
  ch.push(sp(120));
}});

// SECTION 7
ch.push(h1("7. Non-Functional Requirements"));
ch.push(divider());
[
  ["7.1 Performance & Scalability", D.nfrPerf, D.nfrPerfWhy, "b11"],
  ["7.2 Availability & Reliability", D.nfrAvail, D.nfrAvailWhy, "b12"],
  ["7.3 Usability & Accessibility", D.nfrUsab, D.nfrUsabWhy, "b13"],
  ["7.4 Security & Access Control", D.nfrSec, D.nfrSecWhy, "b14"],
  ["7.5 Compliance & Regulatory", D.nfrCompliance, D.nfrComplianceWhy, "b15"],
].forEach(([title, items, why, ref], idx, arr) => {{
  ch.push(h2(title));
  items && items.length > 0
    ? items.forEach(x => ch.push(bullet(x, ref)))
    : ch.push(nullNote(why));
  if (idx < arr.length - 1) ch.push(sp(80));
}});

// SECTION 8
ch.push(h1("8. Data Governance & Privacy"));
ch.push(divider());
ch.push(h2("8.1 Data Classification"));
ch.push(sp(60));
D.dataClassification.length > 0
  ? ch.push(tbl(["Data Asset","Classification Level"],D.dataClassification,[5760,3600]))
  : ch.push(nullNote(D.dataClassificationWhy));
ch.push(sp(160));
ch.push(h2("8.2 Data Privacy Checklist"));
D.privacyChecklist.length > 0
  ? D.privacyChecklist.forEach(x=>ch.push(bullet(x,"b16")))
  : ch.push(nullNote(D.privacyChecklistWhy));

// SECTION 9
ch.push(h1("9. Technology Stack & Architecture"));
ch.push(divider());
ch.push(h2("9.1 Proposed Technology Stack"));
ch.push(body("The development team has been granted autonomy to select the most feasible technology stack."));
ch.push(sp(60));
D.techStack.length > 0
  ? ch.push(tbl(["Architectural Layer","Selected Technology"],D.techStack,[3200,6160]))
  : ch.push(nullNote(D.techStackWhy));
ch.push(sp(200));
ch.push(h2("9.2 Architecture Data Flow"));
ch.push(body("The architecture data flow diagram is to be attached below."));
ch.push(sp(80));
ch.push(archPlaceholder());

// SECTION 10
ch.push(h1("10. Appendices"));
ch.push(divider());
ch.push(h2("10.1 Glossary of Terms"));
ch.push(sp(60));
D.glossary.length > 0
  ? ch.push(tbl(["Term","Definition"],D.glossary,[1800,7560]))
  : ch.push(nullNote(D.glossaryWhy));
ch.push(sp(140));
ch.push(h2("10.2 List of Acronyms"));
D.acronyms.length > 0 ? D.acronyms.forEach(x=>ch.push(bullet(x,"b17"))) : ch.push(nullNote(D.acronymsWhy));
ch.push(sp(120));
ch.push(h2("10.3 Related Documents"));
D.relatedDocs.length > 0 ? D.relatedDocs.forEach(x=>ch.push(bullet(x,"b18"))) : ch.push(nullNote(D.relatedDocsWhy));
ch.push(sp(140));
ch.push(h2("10.4 Document Sign-Off"));
ch.push(sp(60));
const signoffRows = D.signoff.length ? D.signoff : [["________________","________________","________________","________________"]];
ch.push(tbl(["Name","Role","Date","Signature"],signoffRows,[2200,3500,1800,1860]));

const doc = new Document({{
  numbering: {{ config: numberingConfig }},
  styles: {{
    default: {{ document: {{ run: {{ font:"Arial", size:22, color:C.dark }} }} }},
    paragraphStyles: [
      {{ id:"Heading1", name:"Heading 1", basedOn:"Normal", next:"Normal", quickFormat:true,
        run:{{size:36,bold:true,font:"Arial",color:C.primary}},
        paragraph:{{spacing:{{before:480,after:160}},outlineLevel:0}} }},
      {{ id:"Heading2", name:"Heading 2", basedOn:"Normal", next:"Normal", quickFormat:true,
        run:{{size:28,bold:true,font:"Arial",color:C.accent}},
        paragraph:{{spacing:{{before:320,after:120}},outlineLevel:1}} }},
      {{ id:"Heading3", name:"Heading 3", basedOn:"Normal", next:"Normal", quickFormat:true,
        run:{{size:24,bold:true,font:"Arial",color:C.dark}},
        paragraph:{{spacing:{{before:240,after:80}},outlineLevel:2}} }},
    ]
  }},
  sections: [{{
    properties:{{ page:{{size:{{width:12240,height:15840}},margin:{{top:1440,right:1440,bottom:1440,left:1440}}}} }},
    headers:{{ default: new Header({{ children:[
      new Paragraph({{
        children:[
          new TextRun({{text:"CONFIDENTIAL DRAFT | "+D.projectName+" | "+D.businessUnit,size:16,color:C.mid,font:"Arial"}}),
          new TextRun({{text:"\\t",size:16}}),
          new TextRun({{children:["Page ",PageNumber.CURRENT," of ",PageNumber.TOTAL_PAGES],size:16,color:C.mid,font:"Arial"}}),
        ],
        tabStops:[{{type:TabStopType.RIGHT,position:TabStopPosition.MAX}}],
        border:{{bottom:{{style:BorderStyle.SINGLE,size:4,color:C.accent,space:4}}}}
      }})
    ]}}) }},
    footers:{{ default: new Footer({{ children:[
      new Paragraph({{
        children:[new TextRun({{text:"Version "+D.version+" | "+D.date+" | "+D.businessUnit,size:16,color:C.mid,font:"Arial"}})],
        border:{{top:{{style:BorderStyle.SINGLE,size:4,color:C.accent,space:4}}}}
      }})
    ]}}) }},
    children: ch
  }}]
}});

Packer.toBuffer(doc).then(buf => {{
  fs.writeFileSync(process.argv[2], buf);
  console.log("SUCCESS: " + process.argv[2]);
}}).catch(err => {{
  console.error("ERROR: " + err.message);
  process.exit(1);
}});
"""
    return script


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def ensure_docx_package(script_dir: Path):
    if (script_dir / "node_modules" / "docx").exists():
        return
    print("First run: installing 'docx' npm package (~10 seconds)...")
    npm_cmd = "npm.cmd" if sys.platform == "win32" else "npm"
    result = subprocess.run(
        [npm_cmd, "install", "docx", "--prefix", str(script_dir)],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print("ERROR: Failed to install 'docx' npm package.", file=sys.stderr)
        print(result.stderr, file=sys.stderr)
        sys.exit(1)
    print("  'docx' installed successfully.\n")


def main():
    parser = argparse.ArgumentParser(
        description="Generate BRD Word document from structured JSON."
    )
    parser.add_argument("input_json",  help="Path to BRD JSON file")
    parser.add_argument("output_docx", nargs="?",
                        help="Output .docx path (default: <stem>_BRD.docx)")
    args = parser.parse_args()

    input_path = Path(args.input_json).resolve()
    if not input_path.exists():
        print(f"ERROR: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    output_path = (
        Path(args.output_docx).resolve() if args.output_docx
        else input_path.with_name(input_path.stem + "_BRD.docx")
    )

    script_dir = Path(__file__).resolve().parent
    ensure_docx_package(script_dir)

    print(f"Reading JSON: {input_path}")
    with open(input_path, encoding="utf-8") as f:
        data = json.load(f)

    print("Parsing BRD data...")
    ctx = parse_brd(data)
    print(f"  Project      : {ctx['project_name']}")
    print(f"  Phases       : {len(ctx['phases'])}")
    print(f"  Revisions    : {len(ctx['revisions'])}")
    print(f"  Approvals    : {len(ctx['approvals'])}")
    print(f"  Risks        : {len(ctx['risks'])}")
    print(f"  KPIs         : {len(ctx['kpis'])}")
    print(f"  FR sections  : {len(ctx['fr_sections'])}")
    for sec in ctx["fr_sections"]:
        print(f"    {sec['key']}: {len(sec['reqs'])} requirements")
    print(f"  NFR perf     : {len(ctx['nfr_perf'])} items")
    print(f"  NFR avail    : {len(ctx['nfr_avail'])} items")
    print(f"  NFR usab     : {len(ctx['nfr_usab'])} items")
    print(f"  NFR sec      : {len(ctx['nfr_sec'])} items")
    print(f"  NFR compliance: {len(ctx['nfr_compliance'])} items")

    js_code = build_js(ctx)
    tmp_js  = script_dir / "_brd_tmp_gen.js"
    tmp_js.write_text(js_code, encoding="utf-8")

    print(f"Generating document: {output_path}")
    node_cmd = "node.exe" if sys.platform == "win32" else "node"
    result = subprocess.run(
        [node_cmd, str(tmp_js), str(output_path)],
        capture_output=True, text=True, cwd=str(script_dir)
    )
    try:
        tmp_js.unlink()
    except Exception:
        pass

    if result.returncode != 0 or "ERROR" in result.stdout:
        print("Node.js error:", file=sys.stderr)
        print(result.stdout, file=sys.stderr)
        print(result.stderr, file=sys.stderr)
        sys.exit(1)

    print(result.stdout.strip())
    if output_path.exists():
        print(f"\nDone! Output: {output_path} ({output_path.stat().st_size / 1024:.1f} KB)")
    else:
        print("ERROR: Output file was not created.", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
