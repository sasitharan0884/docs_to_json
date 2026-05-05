"""
Structured JSON Builder (Robust, Name+Number Headings)
======================================================

Goal
----
Convert a *lossless* DOCX extraction JSON into a structured JSON that matches the
primary template you shared, without inventing content.

Key guarantees
--------------
- **No generation**: output content is only what exists in the lossless JSON.
- **Robust headings**: detects sections by:
    1) Word heading styles (if present),
    2) numbered headings (e.g., "8.4 ..."),
    3) *named* headings (e.g., "Test Plan", "Tools Required") even when numbers are missing.
- **Section 11** handled as a group: a main "11. Test Execution" container plus
  "11.1.x Test Case Number" subsections; if the main 11 heading is missing,
  it is created as a *container only* when subsections are present (no content is invented).

Input
-----
lossless.json with:
{
  "document": "<docname>.docx",
  "blocks": [
     {"type":"paragraph","text":"...","style":"Heading 1", ...},
     {"type":"table","rows":[...], ...},
     {"type":"image","image_id":"IMG-01","ocr_text":[...], ...},
     ...
  ]
}

Output
------
{
  "document": "...",
  "frontpage_data": {"section_id":"FP-01","content":[...]},
  "sections": [
      {"section_id":"SEC-01","title":"1. ITSAR Section No & Name","level":1,"itsar_section_details":[...]},
      ...
  ]
}

CLI
---
python structured_json_builder.py <lossless.json>
"""

import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


# -------------------------
# Scenario header pattern: only these bold texts start a new scenario
# -------------------------
SCENARIO_HEADER_RE = re.compile(
    r"^\s*test\s*(scenario|case)s?\s+\d+(?:\.\d+){2,}\b",
    re.IGNORECASE
)

# -------------------------
# Helpers
# -------------------------

def _norm(s: str) -> str:
    """Normalize heading-like text for fuzzy-ish matching."""
    s = (s or "").strip().lower()
    s = re.sub(r'[\u2010-\u2015]', '-', s)  # various dashes
    s = re.sub(r'[^a-z0-9\.]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def extract_heading_level(style: str) -> int:
    """Extract heading level from a style string like 'Heading 1', 'heading 2', etc."""
    m = re.search(r'heading\s*(\d+)', style or "", re.IGNORECASE)
    return int(m.group(1)) if m else 1

def normalize_test_case_id(value: str) -> str:
    if not value:
        return value
    cleaned = value.strip()
    cleaned = re.sub(r'[:;.\s]+$', '', cleaned)
    return cleaned

def strip_scenario_header_text(text: str, bold_token: Optional[str]) -> str:
    """Return the body text after a scenario header, handling spacing/punctuation variations."""
    raw = (text or "").strip()
    if not raw:
        return ""

    token = (bold_token or "").strip()
    remainder = raw

    if token and raw.lower().startswith(token.lower()):
        remainder = raw[len(token):].strip()
    else:
        # Fallback to regex match when bold token and text differ slightly.
        m = SCENARIO_HEADER_RE.match(raw)
        if m:
            remainder = raw[m.end():].strip()

    return re.sub(r'^[:\-\s]+', '', remainder).strip()


def strip_leading_bullet_symbol(text: str) -> str:
    """Remove leading bullet/list symbols like '' from tool lines."""
    cleaned = (text or "").strip()
    cleaned = re.sub(r'^\s*[\u2022\u2023\u25E6\u2043\u2219\u25AA\u25CF\u25A0\uF0B7\u00B7]+\s*', '', cleaned)
    cleaned = re.sub(r'^\s*[-*]+\s*', '', cleaned)
    return cleaned.strip()


# -------------------------
# Data structures
# -------------------------

# -------------------------
# Front-page Heading-2 field map (order matters for validation)
# -------------------------
_FP_H2_FIELDS = [
    "DUT Details:",
    "DUT Software Version:",
    "Digest Hash of OS:",
    "Digest Hash of Configuration:",
    "Applicable ITSAR:",
    "ITSAR Version No:",
    "OEM Supplied Document list:",
]

_FP_H2_KEYS = {
    "DUT Details:":                  "dut_details",
    "DUT Software Version:":         "dut_software_version",
    "Digest Hash of OS:":            "digest_hash_of_os",
    "Digest Hash of Configuration:": "digest_hash_of_configuration",
    "Applicable ITSAR:":             "applicable_itsar",
    "ITSAR Version No:":             "itsar_version_no",
    "OEM Supplied Document list:":   "oem_supplied_document_list",
}


def _fp_h2_normalize(text: str) -> str:
    """Normalize a Heading 2 label to lowercase alphanumeric only (e.g. 'digesthashofos')."""
    t = (text or "").lower()
    return re.sub(r'[^a-z0-9]+', '', t)


def validate_frontpage_headings(fp_blocks: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Validate that all required Heading 2 fields are present in the front-page
    blocks and are in the correct order.

    Returns a dict with:
        {
            "status": "PASS" | "FAIL",
            "found_order": [list of matched canonical labels in document order],
            "missing": [list of canonical labels absent from front page],
            "out_of_order": [list of canonical labels found in wrong order],
        }
    """
    found_labels: List[str] = []
    for b in fp_blocks:
        if b.get("type") != "paragraph" or b.get("style") != "Heading 2":
            continue
        raw = _fp_h2_normalize(b.get("text", ""))
        for canonical in _FP_H2_FIELDS:
            norm_can = _fp_h2_normalize(canonical)
            if norm_can == raw or raw.startswith(norm_can):
                found_labels.append(canonical)
                break

    missing = [f for f in _FP_H2_FIELDS if f not in found_labels]

    # Check order: the subsequence of _FP_H2_FIELDS that was found must be in
    # the same relative order as _FP_H2_FIELDS.
    out_of_order: List[str] = []
    expected_order = [f for f in _FP_H2_FIELDS if f in found_labels]
    if found_labels != expected_order:
        out_of_order = [
            f for f in found_labels if f not in expected_order[
                : found_labels.index(f) + 1 if f in expected_order else len(expected_order)
            ]
        ]

    status = "PASS" if not missing and not out_of_order else "FAIL"
    return {
        "status": status,
        "found_order": found_labels,
        "missing": missing,
        "out_of_order": out_of_order,
    }


class FrontPageStructuredExtractor:
    """
    Parses the raw front-page block slice into a typed structured dict.

    Schema produced (replaces the legacy flat-list schema):
    {
      "section_id": "FP-01",
      "content": ["TEST REPORT FOR: ..."],          # title line only
      "dut_details": {
        "title": "DUT Details:",
        "content": [" Product: JIDU6201", " Model: JIDU6201"]
      },
      "dut_software_version": {"title": "DUT Software Version:", "content": ["..."]},
      "digest_hash_of_os":    {"title": "Digest Hash of OS:",    "content": ["..."]},
      "digest_hash_of_configuration": {"title": "Digest Hash of Configuration:", "content": ["..."]},
      "applicable_itsar":     {"title": "Applicable ITSAR:",     "content": ["..."]},
      "itsar_version_no":     {"title": "ITSAR Version No:",     "content": ["..."]},
      "oem_supplied_document_list": {"title": "OEM Supplied Document list:", "content": ["..."]},
    }
    """

    @staticmethod
    def extract(fp_blocks: List[Dict[str, Any]]) -> Dict[str, Any]:
        result: Dict[str, Any] = {
            "section_id": "FP-01",
            "content": [],
        }
        # Initialise all expected field slots
        for canonical, key in _FP_H2_KEYS.items():
            result[key] = {"title": canonical, "content": []}

        current_key: Optional[str] = None

        for b in fp_blocks:
            btype = b.get("type", "")
            style = b.get("style", "")
            text  = (b.get("text") or "").strip()

            if btype != "paragraph":
                continue

            if style == "Heading 2":
                raw_norm = _fp_h2_normalize(text)
                matched = None
                for canonical in _FP_H2_FIELDS:
                    norm_can = _fp_h2_normalize(canonical)
                    if norm_can == raw_norm or raw_norm.startswith(norm_can):
                        matched = canonical
                        break
                if matched:
                    current_key = _FP_H2_KEYS[matched]
                    # Update the title to exactly what's in the document
                    result[current_key]["title"] = text
                else:
                    current_key = None
                continue

            # Non-Heading-2 paragraph
            if current_key is None:
                # Belongs to the overall title/preamble
                if text:
                    result["content"].append(text)
            else:
                if text:
                    result[current_key]["content"].append(text)

        return result


@dataclass
class FrontPage:
    section_id: str = "FP-01"
    content: List[str] = field(default_factory=list)
    structured: Optional[Dict[str, Any]] = None

    def to_dict(self) -> Dict[str, Any]:
        if self.structured is not None:
            return self.structured
        return {"section_id": self.section_id, "content": self.content}

@dataclass
class Section:
    section_id: str
    title: str
    level: int
    content: List[Dict[str, Any]] = field(default_factory=list)
    structured_data: Optional[Dict[str, Any]] = None

    def to_dict(self) -> Dict[str, Any]:
        base = {"section_id": self.section_id, "title": self.title, "level": self.level}
        if self.structured_data is not None:
            base.update(self.structured_data)
        else:
            # Strip internal keys (style, bold_formatting) from content items
            cleaned = []
            for item in self.content:
                if isinstance(item, dict):
                    cleaned.append({k: v for k, v in item.items() if k not in ("style", "bold_formatting")})
                else:
                    cleaned.append(item)
            base["content"] = cleaned
        return base

@dataclass
class DocumentJSON:
    document: str
    frontpage_data: Optional[FrontPage] = None
    sections: List[Section] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "document": self.document,
            "frontpage_data": self.frontpage_data.to_dict() if self.frontpage_data else None,
            "sections": [s.to_dict() for s in self.sections],
        }


# -------------------------
# Section-specific extractors (extract ONLY from section.content)
# -------------------------

class Section1StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict[str, Any]]) -> Dict[str, Any]:
        details = []
        for it in section_content:
            dtype = it.get("type")
            if dtype == "paragraph":
                t = (it.get("text") or "").strip()
                if t:
                    details.append(t)
            elif dtype == "image":
                details.append({"type": "image", "image_path": it.get("image_path", "")})
        return {"itsar_section_details": details}

class Section2StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict[str, Any]]) -> Dict[str, Any]:
        elements = []
        for it in section_content:
            dtype = it.get("type")
            if dtype == "paragraph":
                t = (it.get("text") or "").strip()
                if t:
                    elements.append(t)
            elif dtype == "image":
                elements.append({"type": "image", "image_path": it.get("image_path", "")})
        return {"security_requirement": elements}

class Section3StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict[str, Any]]) -> Dict[str, Any]:
        elements = []
        for it in section_content:
            dtype = it.get("type")
            if dtype == "paragraph":
                t = (it.get("text") or "").strip()
                if t:
                    elements.append(t)
            elif dtype == "image":
                elements.append({"type": "image", "image_path": it.get("image_path", "")})
        return {"requirement_description": elements}

class Section4StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict[str, Any]]) -> Dict[str, Any]:
        dut_details: List[Dict[str, Any]] = []
        for it in section_content:
            t = it.get("type")
            if t == "paragraph":
                txt = (it.get("text") or "").strip()
                if txt:
                    dut_details.append({"type": "paragraph", "text": txt})
            elif t == "table":
                rows = it.get("rows", [])
                if rows and len(rows) >= 2:
                    dut_details.append({
                        "type": "table",
                        "headers": [str(c).strip() for c in rows[0]],
                        "rows": [[str(c).strip() for c in r] for r in rows[1:]],
                    })
            elif t == "image":
                # Keep lossless image info
                dut_details.append({
                    "type": "image",
                    "image_path": it.get("image_path", ""),
                })
        return {"dut_details": dut_details}

class Section5StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict[str, Any]]) -> Dict[str, Any]:
        dut_configuration: List[Dict[str, Any]] = []
        for it in section_content:
            t = it.get("type")
            if t == "paragraph":
                txt = (it.get("text") or "").strip()
                if txt:
                    dut_configuration.append({"type": "paragraph", "text": txt})
            elif t == "table":
                rows = it.get("rows", [])
                if rows and len(rows) >= 2:
                    dut_configuration.append({
                        "type": "table",
                        "headers": [str(c).strip() for c in rows[0]],
                        "rows": [[str(c).strip() for c in r] for r in rows[1:]],
                    })
            elif t == "image":
                dut_configuration.append({
                    "type": "image",
                    "image_path": it.get("image_path", ""),
                })
        return {"dut_configuration": dut_configuration}

class Section6StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict[str, Any]]) -> Dict[str, Any]:
        preconditions = []
        order = 0
        stop_phrases = [
            "test objective", "test plan", "number of test scenarios",
            "tools required", "test execution steps", "expected results",
            "expected format of evidence", "test execution", "test case result",
        ]
        for it in section_content:
            it_type = it.get("type")
            if it_type == "paragraph":
                txt = (it.get("text") or "").strip()
                if not txt:
                    continue
                n = _norm(txt)
                if any(n == _norm(p) or n.startswith(_norm(p)) for p in stop_phrases):
                    break
                preconditions.append({"precondition": txt, "order": order})
                order += 1
            elif it_type == "image":
                preconditions.append({"type": "image", "image_path": it.get("image_path", ""), "order": order})
                order += 1
        return {"preconditions": preconditions, "total_preconditions": len(preconditions)}

class Section81StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict[str, Any]]) -> Dict[str, Any]:
        test_scenarios = []
        current_scenario = None
        current_desc = ""

        def flush():
            nonlocal current_scenario, current_desc
            if current_scenario:
                test_scenarios.append({
                    "test_scenario": current_scenario,
                    "description": current_desc.strip()
                })
            current_scenario = None
            current_desc = ""

        for it in section_content:
            dtype = it.get("type")
            bold = it.get("bold_formatting")
            text = (it.get("text") or "").strip()

            if dtype == "paragraph":
                if not text:
                    continue
                
                if bold and SCENARIO_HEADER_RE.match(text):
                    flush()
                    current_scenario = (bold or "").strip() or text
                    remaining = strip_scenario_header_text(text, bold)
                    if remaining:
                        current_desc = remaining + " "
                    continue
                
                if current_scenario:
                    current_desc += text + " "
            elif dtype == "table" and current_scenario:
                current_desc += f"[Table with {len(it.get('rows', []))} rows] "
            elif dtype == "image" and current_scenario:
                current_desc += f"[Image: {it.get('image_path', '')}] "

        flush()
        return {
            "test_scenarios": test_scenarios,
            "total_test_scenarios": len(test_scenarios)
        }

class Section83StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict[str, Any]]) -> Dict[str, Any]:
        tools = []
        for it in section_content:
            if it.get("type") != "paragraph":
                continue
            txt = (it.get("text") or "").strip()
            if not txt:
                continue
            # stop if semantic boundary
            if re.search(r'\btest\s+execution\s+steps\b', txt, re.IGNORECASE):
                break
            # split comma-separated lists
            parts = [p.strip() for p in txt.split(",") if p.strip()]
            for p in parts:
                cleaned_tool = strip_leading_bullet_symbol(p)
                if cleaned_tool:
                    tools.append({"tool": cleaned_tool})
        
        # Look for images too
        for it in section_content:
            if it.get("type") == "image":
                tools.append({"type": "image", "image_path": it.get("image_path", "")})
        return {"tools": tools, "total_tools": len(tools)}

class Section84StructuredExtractor:
    TEST_SCENARIO_RE = re.compile(r'^Test\s+Scenario\s+(\d+(?:\.\d+)+)', re.IGNORECASE)

    @staticmethod
    def extract(section_content: List[Dict[str, Any]]) -> Dict[str, Any]:
        execution_steps = []
        current_scenario = None
        current_steps: List[Dict[str, Any]] = []
        order = 0

        def flush():
            nonlocal current_scenario, current_steps, order
            if current_scenario:
                execution_steps.append({"test_scenario": current_scenario, "steps": current_steps})
            current_scenario = None
            current_steps = []
            order = 0

        for it in section_content:
            dtype = it.get("type")
            bold = it.get("bold_formatting")

            if dtype == "paragraph":
                txt = (it.get("text") or "").strip()
                if not txt:
                    continue
                
                # Boundary detection: only bold matching scenario header pattern
                if bold and SCENARIO_HEADER_RE.match(txt):
                    flush()
                    current_scenario = (bold or "").strip() or txt
                    remaining = strip_scenario_header_text(txt, bold)
                    if remaining:
                        current_steps.append({"step": remaining, "order": order})
                        order += 1
                    continue
                
                if current_scenario:
                    current_steps.append({"step": txt, "order": order})
                    order += 1
            elif dtype == "table" and current_scenario:
                current_steps.append({"step": "Table", "table": it.get("rows", []), "order": order})
                order += 1
            elif dtype == "image" and current_scenario:
                current_steps.append({"type": "image", "image_path": it.get("image_path", ""), "order": order})
                order += 1

        flush()
        return {"execution_steps": execution_steps, "total_execution_steps": len(execution_steps)}

class Section9StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict[str, Any]]) -> Dict[str, Any]:
        expected_results = []
        current_id = None
        current_text = ""

        def flush():
            nonlocal current_id, current_text
            if current_id:
                expected_results.append({
                    "test_case_id": current_id,
                    "expected_result": current_text.strip()
                })
            current_id = None
            current_text = ""

        for it in section_content:
            dtype = it.get("type")
            bold = it.get("bold_formatting")
            
            if dtype == "paragraph":
                txt = (it.get("text") or "").strip()
                if not txt:
                    continue
                
                if bold and SCENARIO_HEADER_RE.match(txt):
                    flush()
                    current_id = (bold or "").strip() or txt
                    remaining = strip_scenario_header_text(txt, bold)
                    if remaining:
                        current_text = remaining + " "
                    continue
                
                if current_id:
                    current_text += txt + " "
            elif dtype == "image" and current_id:
                current_text += f"[Image: {it.get('image_path', '')}] "

        flush()
        if not expected_results:
            all_text = " ".join(
                (it.get("text") or "").strip()
                for it in section_content
                if it.get("type") == "paragraph" and (it.get("text") or "").strip()
            ).strip()
            if all_text:
                expected_results.append({"test_case_id": "all", "expected_result": all_text})

        return {"expected_results": expected_results, "total_expected_results": len(expected_results)}

class Section12StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict[str, Any]]) -> Dict[str, Any]:
        for item in section_content:
            if item.get("type") != "table":
                continue
            rows = item.get("rows", [])
            if not rows:
                continue
            headers = [str(c).strip() for c in rows[0]]
            body = [[str(c).strip() for c in r] for r in rows[1:]]
            
            # Check for images in Section 12 too
            images = []
            for img in section_content:
                if img.get("type") == "image":
                    images.append({"type": "image", "image_path": img.get("image_path", "")})

            # Standardized headers for JSON output
            standard_headers = ["Test Case Number", "Test Case Name", "Result", "Remarks"]
            final_headers = []
            
            for i, h in enumerate(headers):
                if i < len(standard_headers):
                    final_headers.append(standard_headers[i])
                else:
                    final_headers.append(h)

            return {"test_case_results": {"headers": final_headers, "rows": body, "images": images}}
        return {"test_case_results": {"headers": [], "rows": [], "images": []}}


# -------------------------
# Builder
# -------------------------


# ─────────────────────────────────────────────────────────────────────────────
# Section-11 subsection splitter (for _structured.json only)
# Splits a tc_* raw-block slice into labeled subsection groups.
# ─────────────────────────────────────────────────────────────────────────────

_H3_SLOT_RE = [
    ("test_case_name",        re.compile(r"^a[\s\.\)]", re.I)),
    ("test_case_description", re.compile(r"^b[\s\.\)]", re.I)),
    ("execution_steps",       re.compile(r"^c[\s\.\)]", re.I)),
    ("test_observations",     re.compile(r"^d[\s\.\)]", re.I)),
    ("evidence_provided",     re.compile(r"^e[\s\.\)]", re.I)),
]


def _split_tc_by_subsection(tc_slice: list) -> dict:
    """
    Given a raw tc_* block slice (index 0 = H2 heading block),
    return a dict with these keys (all values are lists of content items):

        test_case_id, test_case_name, test_case_description,
        execution_steps, test_observations, evidence_provided

    H3 heading paragraphs are included as the FIRST item of each group
    (preserving the exact label text, e.g. "a. Test Case Name: ").
    Content before the first H3 goes into test_case_id.
    """
    result = {
        "test_case_id":          [],
        "test_case_name":        [],
        "test_case_description": [],
        "execution_steps":       [],
        "test_observations":     [],
        "evidence_provided":     [],
    }
    current_key = "test_case_id"

    for b in tc_slice[1:]:          # skip the H2 heading block at index 0
        btype = b.get("type", "")
        style = b.get("style", "")
        text  = (b.get("text") or "").strip()

        if btype == "paragraph" and style == "Heading 3":
            # Match against the a-e slot patterns
            matched = None
            for slot_key, pat in _H3_SLOT_RE:
                if pat.match(text):
                    matched = slot_key
                    break
            if matched:
                current_key = matched
                # Include the heading text as the first item of this group
                result[current_key].append({"type": "paragraph", "text": text})
            # Unrecognised H3: ignore (don't change current_key)
            continue

        # Regular content blocks
        if btype == "paragraph":
            full_text = b.get("text", "")
            num_data  = b.get("numbering") or {}
            prefix    = num_data.get("rendered_prefix")
            if prefix and full_text and not full_text.startswith(prefix):
                full_text = f"{prefix} {full_text}"
            if full_text.strip():
                result[current_key].append({"type": "paragraph", "text": full_text})

        elif btype == "table":
            rows = b.get("rows", [])
            if rows and len(rows) >= 2:
                result[current_key].append({
                    "type":    "table",
                    "headers": [str(c).strip() for c in rows[0]],
                    "rows":    [[str(c).strip() for c in r] for r in rows[1:]],
                })
            elif rows:
                result[current_key].append({"type": "table", "rows": rows})

        elif btype == "image":
            img_path = b.get("path", "")
            if img_path:
                result[current_key].append({"type": "image", "image_path": img_path})

    return result

class StructuredSectionBuilder:
    """
    Builds sections in a single pass with robust heading detection.

    Important:
    - We DO NOT fabricate content.
    - We MAY create *empty containers* when Section 11 subsections exist but the
      main "11. Test Execution" heading is missing.
    """

    # Numbered heading (safe): requires whitespace separator after number
    HEADING_RE = re.compile(r'^(\d+(?:\.\d+)*)(?:\.)?\s+(.+)$')

    # Strict expected top-level sections (your primary template order)
    STRICT_SECTIONS = [
        "1. ITSAR Section No & Name",
        "2. Security Requirement No & Name",
        "3. Requirement Description",
        "4. DUT Confirmation Details",
        "5. DUT Configuration",
        "6. Preconditions",
        "7. Test Objective",
        "8. Test Plan",
        "8.1. Number of Test Scenarios",
        "8.2. Test Bed Diagram",
        "8.3. Tools Required",
        "8.4. Test Execution Steps",
        "9. Expected Results for Pass",
        "10. Expected Format of Evidence",
        "11. Test Execution",
        "12. Test Case Result",
    ]

    # Aliases seen across documents (name-only headings, minor variants)
    STRICT_ALIASES = {
        _norm("ITSAR Section No & Name"): "1. ITSAR Section No & Name",
        _norm("Security Requirement No & Name"): "2. Security Requirement No & Name",
        _norm("Requirement Description"): "3. Requirement Description",
        _norm("DUT Confirmation Details"): "4. DUT Confirmation Details",
        _norm("DUT Details"): "4. DUT Confirmation Details",
        _norm("DUT configuration"): "5. DUT Configuration",
        _norm("Preconditions"): "6. Preconditions",
        _norm("Test Objective"): "7. Test Objective",
        _norm("Test Plan"): "8. Test Plan",
        _norm("Number of Test Scenarios"): "8.1. Number of Test Scenarios",
        _norm("Test Bed Diagram"): "8.2. Test Bed Diagram",
        _norm("Tools Required"): "8.3. Tools Required",
        _norm("Test Execution Steps"): "8.4. Test Execution Steps",
        _norm("Expected Results for Pass"): "9. Expected Results for Pass",
        _norm("Expected Results"): "9. Expected Results for Pass",
        _norm("Expected Format of Evidence"): "10. Expected Format of Evidence",
        _norm("Test Execution"): "11. Test Execution",
        _norm("Test Case Result"): "12. Test Case Result",
        _norm("Test Case Results"): "12. Test Case Result",
    }

    # Test case subsection in section 11
    TEST_CASE_SUBSEC_RE = re.compile(r'^(11\.\d+(?:\.\d+)?)\s*\.?\s*Test\s*Case\s*Number\s*:?\s*$', re.IGNORECASE)

    def __init__(self, document_name: str):
        self.output = DocumentJSON(document=document_name)
        self.frontpage = FrontPage()
        self.frontpage_done = False

        self.sections: List[Section] = []
        self.current_section: Optional[Section] = None
        self.section_counter = 0

        # Section 11 tracking
        self.sec11: Optional[Section] = None
        self.in_sec11 = False

    def _next_section_id(self) -> str:
        self.section_counter += 1
        return f"SEC-{self.section_counter:02d}"

    @staticmethod
    def _key_to_section_id(key: str) -> str:
        """Map a split-map key to a deterministic section_id.

        Mapping:
          front_page    → SEC-0
          sec1…sec12   → SEC-1…SEC-12
          sec8_1…sec8_4 → SEC-8-1…SEC-8-4
          tc_N          → SEC-11-N (stamped per-tc inside Section 11)
        """
        if key == "front_page":
            return "SEC-0"
        import re as _re
        m8 = _re.match(r'^sec8_(\d+)$', key)
        if m8:
            return f"SEC-8-{m8.group(1)}"
        m = _re.match(r'^sec(\d+)$', key)
        if m:
            return f"SEC-{m.group(1)}"
        mtc = _re.match(r'^tc_(\d+)$', key)
        if mtc:
            return f"SEC-11-{mtc.group(1)}"
        return f"SEC-{key}"

    def _strict_title_to_level(self, strict_title: str) -> int:
        m = re.match(r'^(\d+(?:\.\d+)*)', strict_title)
        if not m:
            return 1
        return m.group(1).count(".") + 1

    def _match_strict_by_name_or_number(self, text: str) -> Optional[str]:
        """Return canonical strict title if text matches a strict section (numbered or name-only)."""
        t = (text or "").strip()
        if not t:
            return None

        # If it's already a strict section numbered title (or starts with it), map by number
        m = self.HEADING_RE.match(t)
        if m:
            number = m.group(1)
            for s in self.STRICT_SECTIONS:
                if s.startswith(number + ".") or s.startswith(number + " "):
                    return s

        # Match by alias / stripped-number names
        n = _norm(t)
        if n in self.STRICT_ALIASES:
            return self.STRICT_ALIASES[n]

        # Try startswith matching against strict names without numbers
        for s in self.STRICT_SECTIONS:
            s_wo = _norm(re.sub(r'^\d+(?:\.\d+)*\.?\s*', '', s))
            if n == s_wo or n.startswith(s_wo):
                return s

        return None

    def _is_heading(self, block: Dict[str, Any]) -> Tuple[bool, int, str, bool]:
        """
        Returns:
            (is_heading, level, title, is_strict)
        """
        if block.get("type") != "paragraph":
            return False, 0, "", False
        text = (block.get("text") or "").strip()
    def _is_heading(self, blk: Dict[str, Any]) -> Tuple[bool, int, str, bool]:
        """
        Return (is_heading, level, title, is_strict_match)
        Logic:
        - If style == 'Heading 1', it is a Level 1 heading.
        - If style == 'Heading 2', it is a Level 2 heading.
        - Otherwise, strictly ignore.
        """
        if blk.get("type") != "paragraph":
            return False, 0, "", False

        style = blk.get("style", "")
        text = (blk.get("text") or "").strip()
        
        # Check for numbering prefix
        num_data = blk.get("numbering", {})
        prefix = num_data.get("rendered_prefix") if num_data else None

        if style == "Heading 1":
            if prefix and not text.startswith(prefix):
                text = f"{prefix} {text}"
            return True, 1, text, False
        
        if style == "Heading 2":
            if prefix and not text.startswith(prefix):
                text = f"{prefix} {text}"
            return True, 2, text, False

        return False, 0, "", False

    def _section_type(self, title: str) -> str:
        # Use numeric prefix primarily
        m = re.match(r'^(\d+(?:\.\d+)*)', title.strip())
        num = m.group(1) if m else ""
        if num in {"1", "2", "3", "4", "5", "6"}:
            return f"section_{num}"
        if num == "8.1":
            return "section_8_1"
        if num == "8.3":
            return "section_8_3"
        if num == "8.4":
            return "section_8_4"
        if num == "9":
            return "section_9"
        if num == "11":
            return "section_11"
        if num == "12":
            return "section_12"
        return "standard"

    def _apply_extractor(self, section: Section) -> None:
        st = self._section_type(section.title)
        extractor = {
            "section_1": Section1StructuredExtractor,
            "section_2": Section2StructuredExtractor,
            "section_3": Section3StructuredExtractor,
            "section_4": Section4StructuredExtractor,
            "section_5": Section5StructuredExtractor,
            "section_6": Section6StructuredExtractor,
            "section_8_1": Section81StructuredExtractor,
            "section_8_3": Section83StructuredExtractor,
            "section_8_4": Section84StructuredExtractor,
            "section_9": Section9StructuredExtractor,
            "section_12": Section12StructuredExtractor,
        }.get(st)

        if extractor:
            section.structured_data = extractor.extract(section.content)
            section.content = []

    def _ensure_sec11_container(self) -> None:
        if self.sec11 is not None:
            return
        # container only; title is canonical strict title
        self.sec11 = Section(
            section_id=self._next_section_id(),
            title="11. Test Execution",
            level=self._strict_title_to_level("11. Test Execution"),
        )
        self.output.sections.append(self.sec11)

    def build(self, raw_blocks: List[Dict[str, Any]]) -> DocumentJSON:
        i = 0
        while i < len(raw_blocks):
            blk = raw_blocks[i]

            is_head, lvl, title, is_strict = self._is_heading(blk)

            # Frontpage collection until first Heading 1 styled block
            if not self.frontpage_done:
                if is_head:
                    self.frontpage_done = True
                else:
                    if blk.get("type") == "paragraph":
                        t = (blk.get("text") or "").strip()
                        if t:
                            self.frontpage.content.append(t)
                    i += 1
                    continue

            # Special case: absorb requirement IDs into Section 2 if Section 2 active and empty/early
            if is_head and self.current_section and self._section_type(self.current_section.title) == "section_2":
                raw_text = (blk.get("text") or "").strip()
                # requirement-id style like "1.1.2: ..." or "1.1.2 ..."
                if re.match(r'^\d+(\.\d+)+\s*:?', raw_text) and not self._match_strict_by_name_or_number(raw_text):
                    # treat as content
                    self.current_section.content.append({"type": "paragraph", "text": raw_text})
                    i += 1
                    continue

            # Section 11 handling
            if is_head:
                raw_text = (blk.get("text") or "").strip()

                # Start of section 11 by strict title OR by name-only 'Test Execution'
                if title.startswith("11.") and "test execution" in _norm(title) and not self.in_sec11:
                    # close prev
                    if self.current_section:
                        self._apply_extractor(self.current_section)
                    self._ensure_sec11_container()
                    self.current_section = self.sec11
                    self.in_sec11 = True
                    i += 1
                    continue

                # If we encounter 11.x.x Test Case Number, ensure container and create subsection section
                if self.TEST_CASE_SUBSEC_RE.match(raw_text):
                    # close prev (but don't extract for section 11 container)
                    if self.current_section and self.current_section is not self.sec11:
                        self._apply_extractor(self.current_section)
                    self._ensure_sec11_container()
                    self.in_sec11 = True
                    sub = Section(section_id=self._next_section_id(), title=raw_text, level=3)
                    self.output.sections.append(sub)
                    self.current_section = sub
                    i += 1
                    continue

                # Leaving section 11 when we hit Section 12 (by name or number)
                if self.in_sec11 and (title.startswith("12.") or _norm(title) == _norm("12. Test Case Result")):
                    # close current subsection if any
                    if self.current_section and self.current_section is not self.sec11:
                        self._apply_extractor(self.current_section)
                    # close sec11 container (no extractor)
                    self.current_section = None
                    self.in_sec11 = False
                    # proceed normally to create section 12 below

            # Special logic for Heading 2: Only act as a heading if we are in Section 11
            # UNCOMMENTED: To allow global Heading 2 support.
            # if is_head and lvl == 2:
            #     # If we are NOT in section 11 (and not starting it), ignore Heading 2
            #     if not self.in_sec11 and not title.startswith("11."):
            #         is_head = False
            
            # Normal heading -> new section
            if is_head and not (self.TEST_CASE_SUBSEC_RE.match((blk.get("text") or "").strip())):
                # If inside section 11 and heading is another 11.* line that isn't strict, treat as content
                if self.in_sec11 and lvl != 2:
                    # Keep everything under current section 11/subsection unless it's section 12
                    if not (title.startswith("12.") or "test case result" in _norm(title)):
                        # treat as content
                        if blk.get("type") == "paragraph" and (blk.get("text") or "").strip():
                            self.current_section.content.append({
                                "type": "paragraph", 
                                "text": blk["text"], 
                                "style": blk.get("style", ""),
                                "bold_formatting": blk.get("bold_formatting")
                            })
                        i += 1
                        continue

                # close previous section
                if self.current_section:
                    self._apply_extractor(self.current_section)

                # create new
                sec = Section(section_id=self._next_section_id(), title=title, level=lvl)
                self.output.sections.append(sec)
                self.current_section = sec

                # update section 11 tracking
                self.in_sec11 = title.startswith("11.")
                if self.in_sec11:
                    self.sec11 = sec

                i += 1
                continue

            # Add content to current section
            if self.current_section is not None:
                t = blk.get("type")
                if t == "paragraph":
                    txt = (blk.get("text") or "").strip()
                    if txt:
                        # Prepend prefix if present (e.g. for bullets)
                        num_data = blk.get("numbering", {})
                        prefix = num_data.get("rendered_prefix") if num_data else None
                        full_text = blk.get("text")
                        if prefix and not full_text.startswith(prefix):
                            full_text = f"{prefix} {full_text}"
                            
                        self.current_section.content.append({
                            "type": "paragraph", 
                            "text": full_text,
                            "style": blk.get("style", ""),
                            "bold_formatting": blk.get("bold_formatting")
                        })
                elif t == "table":
                    self.current_section.content.append({"type": "table", "rows": blk.get("rows", [])})
                elif t == "image":
                    self.current_section.content.append({
                        "type": "image",
                        "image_path": blk.get("path", ""),
                    })
                else:
                    # preserve unknown blocks
                    self.current_section.content.append(blk)

            i += 1

        # finalize last section
        if self.current_section:
            self._apply_extractor(self.current_section)

        self.output.frontpage_data = self.frontpage
        return self.output

    # ------------------------------------------------------------------
    # NEW: Map-driven build — validator owns structure, extractor owns content
    # ------------------------------------------------------------------
    def build_from_map(
        self,
        mapped_sections: Dict[str, List[Dict[str, Any]]],
        failed_keys: Optional[set] = None,
        failed_tc_headings: Optional[Dict[str, str]] = None,
    ) -> "DocumentJSON":
        """
        Build structured output directly from pre-sliced mapped_sections produced
        by document_split_validator.build_mapped_sections().

        Each mapped_sections[key] is a list of raw lossless blocks for that section.
        Index [0] is the heading block; content starts at index [1].

        Benefits over build(all_blocks):
          - No heading re-detection (validator already owns structure).
          - No section-bleed (each slice is bounded).
          - Only iterates over relevant blocks per section (faster).
          - Single source of truth: validator → extractor.
        """
        # Section key → (canonical title, level)
        SEC_META: Dict[str, Tuple[str, int]] = {
            "front_page": ("[Front Page]", 0),
            "sec1":   ("1. ITSAR Section No & Name",        1),
            "sec2":   ("2. Security Requirement No & Name", 1),
            "sec3":   ("3. Requirement Description",        1),
            "sec4":   ("4. DUT Confirmation Details",       1),
            "sec5":   ("5. DUT Configuration",              1),
            "sec6":   ("6. Preconditions",                  1),
            "sec7":   ("7. Test Objective",                 1),
            "sec8":   ("8. Test Plan",                      1),
            "sec8_1": ("8.1. Number of Test Scenarios",     2),
            "sec8_2": ("8.2. Test Bed Diagram",             2),
            "sec8_3": ("8.3. Tools Required",               2),
            "sec8_4": ("8.4. Test Execution Steps",         2),
            "sec9":   ("9. Expected Results for Pass",      1),
            "sec10":  ("10. Expected Format of Evidence",   1),
            "sec11":  ("11. Test Execution",                1),
            "sec12":  ("12. Test Case Result",              1),
        }

        def _to_content(raw_slice: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
            """Convert raw block slice (skip [0] heading) into extractor-ready content."""
            content: List[Dict[str, Any]] = []
            for b in raw_slice[1:]:
                btype = b.get("type", "")
                if btype == "paragraph":
                    txt = (b.get("text") or "").strip()
                    num_data = b.get("numbering", {})
                    prefix = (num_data or {}).get("rendered_prefix")
                    full_text = b.get("text", "")
                    if prefix and full_text and not full_text.startswith(prefix):
                        full_text = f"{prefix} {full_text}"
                    if full_text.strip():
                        content.append({
                            "type": "paragraph",
                            "text": full_text,
                            "style": b.get("style", ""),
                            "bold_formatting": b.get("bold_formatting"),
                        })
                elif btype == "table":
                    content.append({"type": "table", "rows": b.get("rows", [])})
                elif btype == "image":
                    content.append({"type": "image", "image_path": b.get("path", "")})
            return content

        # Front-page — validate Heading 2 order, then build structured output
        fp_blocks = mapped_sections.get("front_page", [])
        _fp_validation = validate_frontpage_headings(fp_blocks)
        if _fp_validation["status"] == "PASS":
            self.frontpage.structured = FrontPageStructuredExtractor.extract(fp_blocks)
        else:
            # One or more required Heading 2 fields are missing — emit a FAIL stub.
            # The full error detail is surfaced in output.json (section_0).
            self.frontpage.structured = {
                "section_id": "FP-01",
                "status":     "FAIL",
            }
        self.frontpage_done = True
        self.output.frontpage_data = self.frontpage

        ORDERED_KEYS = [
            "sec1", "sec2", "sec3", "sec4", "sec5", "sec6", "sec7",
            "sec8", "sec8_1", "sec8_2", "sec8_3", "sec8_4",
            "sec9", "sec10", "sec11", "sec12",
        ]

        # Collect tc_* keys for Section 11
        tc_keys = sorted(
            [k for k in mapped_sections if k.startswith("tc_")],
            key=lambda k: int(k.split("_")[1]),
        )

        failed_keys = failed_keys or set()

        for key in ORDERED_KEYS:
            raw_slice = mapped_sections.get(key)
            if raw_slice is None:
                # Emit a FAIL stub if this key was explicitly failed
                if key in failed_keys and key in SEC_META:
                    canonical_title, level = SEC_META[key]
                    stub = Section(
                        section_id=self._key_to_section_id(key),
                        title=canonical_title,
                        level=level,
                    )
                    stub.structured_data = {"status": "FAIL"}
                    self.output.sections.append(stub)
                continue

            canonical_title, level = SEC_META.get(key, (key, 1))
            content = _to_content(raw_slice)

            sec = Section(
                section_id=self._key_to_section_id(key),
                title=canonical_title,
                level=level,
            )

            st = self._section_type(canonical_title)

            if st == "section_8_1":
                sec.structured_data = Section81StructuredExtractor.extract(content)
            elif st == "section_8_3":
                sec.structured_data = Section83StructuredExtractor.extract(content)
            elif st == "section_8_4":
                sec.structured_data = Section84StructuredExtractor.extract(content)
            elif st == "section_9":
                sec.structured_data = Section9StructuredExtractor.extract(content)
            elif st == "section_12":
                sec.structured_data = Section12StructuredExtractor.extract(content)
            elif st in ("section_1", "section_2", "section_3", "section_4",
                        "section_5", "section_6"):
                extractor_cls = {
                    "section_1": Section1StructuredExtractor,
                    "section_2": Section2StructuredExtractor,
                    "section_3": Section3StructuredExtractor,
                    "section_4": Section4StructuredExtractor,
                    "section_5": Section5StructuredExtractor,
                    "section_6": Section6StructuredExtractor,
                }.get(st)
                if extractor_cls:
                    sec.structured_data = extractor_cls.extract(content)
            elif st == "section_11":
                # Section 11 container: add preamble content (if any) then skip;
                # each tc_* slice becomes its own flat subsection below.
                sec.content = content  # preamble paragraphs (usually empty)
                self.output.sections.append(sec)

                # Build the full ordered tc_* list: merge passing keys (from mapped_sections)
                # and failing keys (from failed_tc_headings) sorted by their numeric suffix,
                # so output appears as tc_1, tc_2, tc_3... regardless of PASS/FAIL.
                _failed_tc_hdgs = failed_tc_headings or {}
                _all_tc_keys = sorted(
                    list(tc_keys) + [k for k in _failed_tc_hdgs],
                    key=lambda k: int(k.split("_")[1]) if k.split("_")[1].isdigit() else 99999,
                )

                for tc_key in _all_tc_keys:
                    tc_num = tc_key.split("_")[-1]

                    if tc_key in tc_keys:
                        # PASS TC: split content into labeled subsection groups.
                        tc_slice = mapped_sections.get(tc_key, [])
                        if not tc_slice:
                            continue

                        tc_title = (tc_slice[0].get("text") or "").strip()

                        tc_sec = Section(
                            section_id="SEC-11-" + tc_num,
                            title=tc_title,
                            level=3,
                        )

                        # Organise content by H3 subsection labels
                        # (test_case_id, test_case_name, test_case_description,
                        #  execution_steps, test_observations, evidence_provided)
                        tc_sec.structured_data = _split_tc_by_subsection(tc_slice)
                        self.output.sections.append(tc_sec)

                    else:
                        # FAIL TC: emit a minimal stub.
                        # Title = exact H2 text from the document (e.g. "11.1.7 Test Case Number:")
                        actual_heading = _failed_tc_hdgs.get(tc_key, tc_key)
                        fail_stub = Section(
                            section_id="SEC-11-" + tc_num,
                            title=actual_heading,
                            level=3,
                        )
                        fail_stub.structured_data = {"status": "FAIL"}
                        self.output.sections.append(fail_stub)

                # Skip the generic append at the bottom of the loop since we already appended
                continue
            else:
                sec.content = content

            self.output.sections.append(sec)

        return self.output


# -------------------------
# Public functions / CLI
# -------------------------

def build_structured_document(lossless_json_path: str) -> Dict[str, Any]:
    with open(lossless_json_path, "r", encoding="utf-8") as f:
        lossless = json.load(f)
    builder = StructuredSectionBuilder(lossless.get("document", "document.docx"))
    doc = builder.build(lossless.get("blocks", []))
    return doc.to_dict()

def main():
    if len(sys.argv) < 2:
        print("Usage: python structured_json_builder.py <lossless.json>")
        sys.exit(1)
    lossless_path = Path(sys.argv[1])
    if not lossless_path.exists():
        raise SystemExit(f"File not found: {lossless_path}")
    structured = build_structured_document(str(lossless_path))
    out_path = lossless_path.parent / f"{lossless_path.stem.replace('_lossless','')}_structured.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(structured, f, indent=2, ensure_ascii=False)
    print(f"✅ STRUCTURED DOCUMENT CREATED: {out_path}")

if __name__ == "__main__":
    main()
