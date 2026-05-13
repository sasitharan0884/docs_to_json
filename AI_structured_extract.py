"""
AI Structured JSON Builder (Clean Single-File Version)
------------------------------------------------------
This is the cleaned, final AI_structured_extract.py with:
- ONE StructuredSectionBuilder (no duplicates)
- Section 8.1 / 8.4 / 9 scenario boundaries driven by bold_formatting
- Section 11 uses Heading 2 as test case boundary and Heading 3 as subsections
- Section 12 parses results table
- Other sections remain structured as content lists

IMPORTANT ENFORCED FORMAT (per your rule):
- Section 8.1: only scenario token is bold (bold_formatting contains only that token)
- Section 8.4: entire scenario header line is bold (bold_formatting contains full header line)
- Section 9: only scenario token is bold (bold_formatting contains only that token)
"""

import json
import sys
import re
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple


# -------------------------------------------------
# Scenario header pattern: only these bold texts start a new scenario
# -------------------------------------------------
SCENARIO_HEADER_RE = re.compile(
    r"^\s*test\s*(scenario|case)s?\s+\d+(?:\.\d+){2,}\b",
    re.IGNORECASE
)

# Matches bare "Test Scenario: ..." or "Test Case: ..." with no dotted number.
# Used so Sections 8.1, 8.4, and 9 all recognise the same header style.
_SCENARIO_BARE_RE = re.compile(
    r"^\s*test\s*(?:scenario|case)s?\s*:\s*",
    re.IGNORECASE,
)


def _is_bold_scenario_header(text: str, bold: str) -> bool:
    """Return True if *text* opens a new scenario block and is formatted bold.

    Handles both:
    - Numbered: "Test Scenario 1.9.2.1: ..."
    - Bare label: "Test Scenario: Verify mutual Authentication."
    """
    if not bold:
        return False
    if SCENARIO_HEADER_RE.match(text):
        return True
    if _SCENARIO_BARE_RE.match(text):
        return True
    return False

# -------------------------------------------------
# NAME-ONLY NORMALISATION HELPERS
# -------------------------------------------------
NUM_PREFIX_RE = re.compile(r"^\s*\d+(?:\.\d+)*\s*[\.):]?\s*")


def strip_num_prefix(text: str) -> str:
    """Remove leading numbering like '8.', '7.', '8.1.' from text."""
    return NUM_PREFIX_RE.sub("", (text or "").strip())


def norm_name(text: str) -> str:
    """Normalize section name only — ignore numbering and punctuation."""
    t = strip_num_prefix(text)
    t = t.strip().lower()
    t = re.sub(r"[\u2010-\u2015]", "-", t)   # normalize dashes
    t = re.sub(r"[^a-z0-9]+", " ", t)        # keep only alphanum
    return re.sub(r"\s+", " ", t).strip()


def normalize_test_case_id(value: str) -> str:
    """Normalize test case ID by stripping trailing punctuation and whitespace."""
    if not value:
        return value
    cleaned = value.strip()
    cleaned = re.sub(r'[:;.\s]+$', '', cleaned)
    return cleaned


def extract_scenario_id(bold: str, text: str) -> str:
    """
    Extract the full scenario label (e.g. 'Test Scenario 1.9.2.1:') 
    without the subsequent descriptive text.
    """
    raw = (bold or "").strip() or text.strip()
    
    # 1. Try to match the standard numbered prefix
    m = SCENARIO_HEADER_RE.match(raw)
    if m:
        prefix = m.group(0).strip()
        # Check if a colon immediately follows the match
        after_match = raw[m.end():]
        if after_match.startswith(":"):
            return prefix + ":"
        return prefix

    # 2. Try bare label prefix
    m_bare = _SCENARIO_BARE_RE.match(raw)
    if m_bare:
        return m_bare.group(0).strip()

    # Fallback: cleaned version of the start of the line
    token = (bold or "").strip()
    if token and len(token) < 50:
        return token
    
    # Very aggressive fallback: take text before first colon
    if ":" in raw:
        return raw.split(":")[0].strip() + ":"
    return raw.strip()


def strip_scenario_header_text(text: str, bold_token: Optional[str]) -> str:
    """Return content after scenario header while tolerating formatting differences."""
    raw = (text or "").strip()
    if not raw:
        return ""

    # 1. Try numeric ID boundary first (matches "Test Scenario 1.2.3.4")
    m = SCENARIO_HEADER_RE.match(raw)
    if m:
        remainder = raw[m.end():].strip()
        return re.sub(r'^[:\-\s]+', '', remainder).strip()

    # 2. Try bare label boundary (matches "Test Scenario:")
    m_bare = _SCENARIO_BARE_RE.match(raw)
    if m_bare:
        remainder = raw[m_bare.end():].strip()
        return re.sub(r'^[:\-\s]+', '', remainder).strip()

    # 3. Fallback to bold token stripping
    token = (bold_token or "").strip()
    if token and raw.lower().startswith(token.lower()):
        remainder = raw[len(token):].strip()
        return re.sub(r'^[:\-\s]+', '', remainder).strip()

    return ""


def strip_leading_bullet_symbol(text: str) -> str:
    """Remove leading bullet/list symbols like '' from tool lines."""
    cleaned = (text or "").strip()
    cleaned = re.sub(r'^\s*[\u2022\u2023\u25E6\u2043\u2219\u25AA\u25CF\u25A0\uF0B7\u00B7]+\s*', '', cleaned)
    cleaned = re.sub(r'^\s*[-*]+\s*', '', cleaned)
    return cleaned.strip()


# -------------------------------------------------
# Front-page Heading-2 field map (order matters for validation)
# -------------------------------------------------
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
      "content": ["TEST REPORT FOR: ..."],
      "dut_details":                  {"title": "DUT Details:",                  "content": [...]},
      "dut_software_version":         {"title": "DUT Software Version:",         "content": [...]},
      "digest_hash_of_os":            {"title": "Digest Hash of OS:",            "content": [...]},
      "digest_hash_of_configuration": {"title": "Digest Hash of Configuration:", "content": [...]},
      "applicable_itsar":             {"title": "Applicable ITSAR:",             "content": [...]},
      "itsar_version_no":             {"title": "ITSAR Version No:",             "content": [...]},
      "oem_supplied_document_list":   {"title": "OEM Supplied Document list:",   "content": [...]},
    }
    """

    @staticmethod
    def extract(fp_blocks: List[Dict[str, Any]]) -> Dict[str, Any]:
        result: Dict[str, Any] = {"section_id": "FP-01", "content": []}
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
                    result[current_key]["title"] = text
                else:
                    current_key = None
                continue

            if current_key is None:
                if text:
                    result["content"].append(text)
            else:
                if text:
                    result[current_key]["content"].append(text)

        return result


# -------------------------------------------------
# DATA STRUCTURES
# -------------------------------------------------
class FrontPage:
    def __init__(self):
        self.section_id = "FP-01"
        self.content: List[Any] = []
        self.structured: Optional[Dict[str, Any]] = None

    def to_dict(self):
        if self.structured is not None:
            return self.structured
        return {"section_id": self.section_id, "content": self.content}


class Section:
    def __init__(self, section_id: str, title: str, level: int):
        self.section_id = section_id
        self.title = title
        self.level = level
        self.content: List[Dict[str, Any]] = []
        self.structured_data: Optional[Dict[str, Any]] = None
        self.extracted: bool = False
    def to_dict(self):
        out = {"section_id": self.section_id, "title": self.title, "level": self.level}
        if self.structured_data is not None:
            out.update(self.structured_data)
        else:
            # Strip internal keys (style, bold_formatting) from content items
            cleaned = []
            for item in self.content:
                if isinstance(item, dict):
                    cleaned.append({k: v for k, v in item.items() if k not in ("style", "bold_formatting")})
                else:
                    cleaned.append(item)
            out["content"] = cleaned
        return out


class DocumentJSON:
    def __init__(self, document_name: str):
        self.document = document_name
        self.frontpage_data: Optional[FrontPage] = None
        self.sections: List[Section] = []

    def to_dict(self):
        return {
            "document": self.document,
            "frontpage_data": self.frontpage_data.to_dict() if self.frontpage_data else None,
            "sections": [s.to_dict() for s in self.sections],
        }


# -------------------------------------------------
# STRUCTURED EXTRACTORS
# -------------------------------------------------
class Section1StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict]) -> Dict[str, Any]:
        details = []
        for item in section_content:
            if not isinstance(item, dict):
                continue
            itype = item.get("type")
            if itype == "paragraph":
                text = (item.get("text") or "").strip()
                if text:
                    details.append(text)
            elif itype == "image":
                details.append({"type": "image", "image_path": item.get("image_path", "")})
        return {"itsar_section_details": details}


class Section2StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict]) -> Dict[str, Any]:
        elements = []
        for item in section_content:
            if not isinstance(item, dict):
                continue
            itype = item.get("type")
            if itype == "paragraph":
                text = (item.get("text") or "").strip()
                if text:
                    elements.append(text)
            elif itype == "image":
                elements.append({"type": "image", "image_path": item.get("image_path", "")})
        return {"security_requirement": elements}


class Section3StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict]) -> Dict[str, Any]:
        elements = []
        for item in section_content:
            if not isinstance(item, dict):
                continue
            itype = item.get("type")
            if itype == "paragraph":
                text = (item.get("text") or "").strip()
                if text:
                    elements.append(text)
            elif itype == "image":
                elements.append({"type": "image", "image_path": item.get("image_path", "")})
        return {"requirement_description": elements}


class Section4StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict]) -> Dict[str, Any]:
        dut_details = []
        for item in section_content:
            if not isinstance(item, dict):
                continue
            item_type = item.get("type")
            if item_type == "paragraph":
                text = (item.get("text") or "").strip()
                if text:
                    dut_details.append({"type": "paragraph", "text": text})
            elif item_type == "table":
                rows = item.get("rows", [])
                if not rows or len(rows) < 2:
                    continue
                dut_details.append({
                    "type": "table",
                    "headers": [cell.strip() for cell in rows[0]],
                    "rows": [[cell.strip() for cell in row] for row in rows[1:]],
                })
            elif item_type == "image":
                dut_details.append({"type": "image", "image_path": item.get("image_path", "")})
        return {"dut_details": dut_details}


class Section5StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict]) -> Dict[str, Any]:
        dut_configuration = []
        for item in section_content:
            if not isinstance(item, dict):
                continue
            item_type = item.get("type")
            if item_type == "paragraph":
                text = (item.get("text") or "").strip()
                if text:
                    dut_configuration.append({"type": "paragraph", "text": text})
            elif item_type == "table":
                rows = item.get("rows", [])
                if not rows or len(rows) < 2:
                    # If it's a simple table, maybe just rows? 
                    # But Section 4 style uses headers/rows.
                    if rows:
                        dut_configuration.append({
                            "type": "table",
                            "rows": [[cell.strip() if cell else "" for cell in row] for row in rows]
                        })
                    continue
                dut_configuration.append({
                    "type": "table",
                    "headers": [str(cell).strip() for cell in rows[0]],
                    "rows": [[str(cell).strip() for cell in row] for row in rows[1:]],
                })
            elif item_type == "image":
                dut_configuration.append({"type": "image", "image_path": item.get("image_path", "")})
        return {"dut_configuration": dut_configuration}


class Section6StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict]) -> Dict[str, Any]:
        preconditions = []
        order = 0
        for item in section_content:
            if not isinstance(item, dict):
                continue
            itype = item.get("type")
            if itype == "paragraph":
                text = (item.get("text") or "").strip()
                if text:
                    preconditions.append({"precondition": text, "order": order})
                    order += 1
            elif itype == "image":
                preconditions.append({"type": "image", "image_path": item.get("image_path", ""), "order": order})
                order += 1
        return {"preconditions": preconditions, "total_preconditions": len(preconditions)}


class Section83StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict]) -> Dict[str, Any]:
        tools = []
        for item in section_content:
            if not isinstance(item, dict):
                continue
            itype = item.get("type")
            if itype == "paragraph":
                text = (item.get("text") or "").strip()
                if not text:
                    continue
                cleaned_tool = strip_leading_bullet_symbol(text)
                if cleaned_tool:
                    tools.append({"tool": cleaned_tool})
            elif itype == "image":
                tools.append({"type": "image", "image_path": item.get("image_path", "")})
        return {"tools": tools, "total_tools": len([t for t in tools if isinstance(t, dict) and "tool" in t])}


class Section81StructuredExtractor:
    """
    Section 8.1:
    - A new scenario begins when bold_formatting is present (token-only bold)
    - Remaining text (after removing bold token) belongs to description
    """
    @staticmethod
    def extract(section_content: List[Dict]) -> Dict[str, Any]:
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

        for item in section_content:
            if not isinstance(item, dict):
                continue
            itype = item.get("type", "")
            bold = (item.get("bold_formatting") or "").strip()
            text = (item.get("text") or "").strip()

            if itype == "paragraph":
                if not text:
                    continue
                if _is_bold_scenario_header(text, bold):
                    flush()
                    current_scenario = extract_scenario_id(bold, text)
                    remaining = strip_scenario_header_text(text, bold)
                    if remaining:
                        current_desc = remaining + " "
                    continue
                if current_scenario:
                    current_desc += text + " "
            elif itype == "table" and current_scenario:
                current_desc += f"[Table with {len(item.get('rows', []))} rows] "
            elif itype == "image" and current_scenario:
                current_desc += f"[Image: {item.get('image_path', '')}] "

        flush()
        return {"test_scenarios": test_scenarios, "total_test_scenarios": len(test_scenarios)}


class Section84StructuredExtractor:
    """
    Section 8.4:
    - Entire scenario header line is bold (bold_formatting contains full header)
    - New scenario when bold_formatting exists
    - Subsequent items become steps until next bold header
    """
    @staticmethod
    def extract(section_content: List[Dict]) -> Dict[str, Any]:
        execution_steps = []
        current_scenario = None
        current_steps = []
        order = 0

        def flush():
            nonlocal current_scenario, current_steps, order
            if current_scenario:
                execution_steps.append({"test_scenario": current_scenario, "steps": current_steps})
            current_scenario = None
            current_steps = []
            order = 0

        for item in section_content:
            if not isinstance(item, dict):
                continue

            itype = item.get("type", "")
            bold = (item.get("bold_formatting") or "").strip()
            text = (item.get("text") or "").strip()

            if itype == "paragraph":
                if not text:
                    continue
                if _is_bold_scenario_header(text, bold):
                    flush()
                    current_scenario = extract_scenario_id(bold, text)
                    remaining = strip_scenario_header_text(text, bold)
                    if remaining:
                        current_steps.append({"step": remaining, "order": order})
                        order += 1
                    continue
                if current_scenario:
                    current_steps.append({"step": text, "order": order})
                    order += 1

            elif itype == "table" and current_scenario:
                current_steps.append({"step": "Table", "table": item.get("rows", []), "order": order})
                order += 1

            elif itype == "image" and current_scenario:
                current_steps.append({"type": "image", "image_path": item.get("image_path", ""), "order": order})
                order += 1

        flush()
        return {"execution_steps": execution_steps, "total_execution_steps": len(execution_steps)}


class Section9StructuredExtractor:
    """
    Section 9:
    - New expected-result entry when bold_formatting exists (token-only bold)
    - Remaining text (after removing bold token) belongs to expected_result

    Guards:
    - If any Heading 2 or Heading 3 paragraph is present in section_content
      the section boundary is contaminated — returns an error stub and no results.

    Scenario ID matching (broader than module-level SCENARIO_HEADER_RE):
    - "Test Scenario 1.1.11"  (two-segment dotted suffix)
    - "Test Scenario:"        (bare label at start of text)
    """

    # Matches "Test Scenario" or "Test Case" followed by a dotted number with
    # at least ONE dot (e.g. 1.1, 1.1.7, 1.1.11) — more permissive than the
    # module-level SCENARIO_HEADER_RE which requires {2,} dot groups.
    _SCENARIO_ID_RE = re.compile(
        r"^\s*test\s*(?:scenario|case)s?\s+\d+(?:\.\d+){1,}\b",
        re.IGNORECASE,
    )

    # Matches a bare "Test Scenario:" / "Test Case:" label with no number.
    _SCENARIO_BARE_RE = re.compile(
        r"^\s*test\s*(?:scenario|case)s?\s*:\s*",
        re.IGNORECASE,
    )

    @classmethod
    def _is_scenario_header(cls, text: str, bold: str) -> bool:
        """Return True if this paragraph opens a new expected-result entry."""
        # Standard: text starts with a dotted test-scenario ID and is bold
        if bold and cls._SCENARIO_ID_RE.match(text):
            return True
        # Bare label "Test Scenario: ..." — bold required to avoid false positives
        if bold and cls._SCENARIO_BARE_RE.match(text):
            return True
        return False

    @classmethod
    def extract(cls, section_content: List[Dict]) -> Dict[str, Any]:
        # ---------------------------------------------------------------
        # Guard: Heading 2 / Heading 3 present → boundary contamination
        # ---------------------------------------------------------------
        heading_issues = []
        for item in section_content:
            if not isinstance(item, dict):
                continue
            style = (item.get("style") or "").strip()
            if style in ("Heading 2", "Heading 3"):
                heading_issues.append({
                    "type": "UNEXPECTED_HEADING",
                    "severity": "HIGH",
                    "style": style,
                    "text": (item.get("text") or "").strip(),
                    "message": (
                        f"A '{style}' paragraph was found inside Section 9 "
                        "(Expected Results for Pass). This indicates a boundary "
                        "issue — test-case sub-headings should not appear here."
                    ),
                })

        if heading_issues:
            return {
                "status": "FAIL",
                "error": "Section 9 contains unexpected Heading 2 / Heading 3 paragraphs. "
                         "Expected-result extraction was skipped to avoid contaminated output.",
                "heading_issues": heading_issues,
                "expected_results": [],
                "total_expected_results": 0,
            }

        # ---------------------------------------------------------------
        # Normal extraction
        # ---------------------------------------------------------------
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

        for item in section_content:
            if not isinstance(item, dict):
                continue

            itype = item.get("type", "")
            bold = (item.get("bold_formatting") or "").strip()
            text = (item.get("text") or "").strip()

            if itype == "paragraph":
                if not text:
                    continue
                if cls._is_scenario_header(text, bold):
                    flush()
                    current_id = extract_scenario_id(bold, text)
                    remaining = strip_scenario_header_text(text, bold)
                    current_text = (remaining + " ") if remaining else ""
                    continue
                if current_id:
                    current_text += text + " "

            elif itype == "image" and current_id:
                current_text += f"[Image: {item.get('image_path', '')}] "

        flush()

        if not expected_results:
            all_text = " ".join(
                (it.get("text") or "").strip()
                for it in section_content
                if isinstance(it, dict) and it.get("type") == "paragraph"
            ).strip()
            if all_text:
                expected_results.append({"test_case_id": "all", "expected_result": all_text})

        return {
            "expected_results": expected_results,
            "total_expected_results": len([r for r in expected_results if "test_case_id" in r])
        }


class Section12StructuredExtractor:
    @staticmethod
    def extract(section_content: List[Dict]) -> Dict[str, Any]:
        test_results = []
        for item in section_content:
            if not isinstance(item, dict):
                continue
            if item.get("type") == "table":
                rows = item.get("rows", [])
                for row in rows[1:] if len(rows) > 1 else []:
                    if len(row) >= 4:
                        # row[0] = test case ID/number, row[1] = name, row[2] = result, row[3] = remarks
                        test_case_id  = normalize_test_case_id((row[0] or "").strip())
                        test_case_name = (row[1] or "").strip()
                        result_status  = re.sub(r"\s+", " ", (row[2] or "").strip())
                        remarks        = (row[3] or "").strip()
                        if test_case_id or test_case_name:
                            test_results.append({
                                "test_case_id":   test_case_id,
                                "test_case_name": test_case_name,
                                "result":         result_status,
                                "remarks":        remarks,
                            })
            elif item.get("type") == "image":
                test_results.append({"type": "image", "image_path": item.get("image_path", "")})

        return {
            "test_results": test_results,
            "total_results": len([r for r in test_results if isinstance(r, dict) and r.get("test_case_id")])
        }


class Section11StructuredExtractor:
    """Heading 2 = test case title, Heading 3 = a-e subsections"""

    _TC_PATTERNS = [
        re.compile(r'TC-(\d+(?:\.\d+){3})'),
        re.compile(r'(\d+(?:\.\d+){3})'),
    ]

    # Regex to match Heading 3 labels (a-e) and their keywords to allow splitting even without colons.
    _H3_LABEL_RE = re.compile(
        r'^(([a-e][\s\.\)]\s*)?(test\s*case\s*)?(name|description|execution|observation|evidence)[\w\s]*?[\s\.\-\:]+)',
        re.IGNORECASE
    )

    @classmethod
    def _extract_tc_id(cls, text: str) -> str:
        for pat in cls._TC_PATTERNS:
            m = pat.search(text or "")
            if m:
                return normalize_test_case_id(m.group(0))
        return ""

    @staticmethod
    def _map_heading3(text: str) -> Optional[str]:
        t = (text or "").strip().lower()
        if re.match(r'^a[\s\.\)]', t) and 'test case name' in t:
            return 'name'
        if re.match(r'^b[\s\.\)]', t) and 'test case description' in t:
            return 'description'
        if re.match(r'^c[\s\.\)]', t) and 'execution' in t:
            return 'execution'
        if re.match(r'^d[\s\.\)]', t) and 'observation' in t:
            return 'observation'
        if re.match(r'^e[\s\.\)]', t) and 'evidence' in t:
            return 'evidence'
        if 'test case name' in t and 'description' not in t:
            return 'name'
        if 'test case description' in t:
            return 'description'
        if 'execution' in t:
            return 'execution'
        if 'observation' in t:
            return 'observation'
        if 'evidence' in t:
            return 'evidence'
        return None

    @classmethod
    def extract(cls, section_content: List[Dict]) -> Dict[str, Any]:
        test_cases: List[Dict[str, Any]] = []
        current_tc: Optional[Dict[str, Any]] = None
        current_sub: Optional[str] = None
        exec_order = 0
        evidence_order = 0

        # All 5 subsection slots that must be present as Heading 3
        REQUIRED_H3S = {"name", "description", "execution", "observation", "evidence"}

        def new_test_case(heading_text: str) -> Dict[str, Any]:
            return {
                "test_case_heading": heading_text.strip(),
                "test_case_id": cls._extract_tc_id(heading_text),
                "test_case_name": "",
                "test_case_description": "",
                "execution": [],
                "test_observation": [],
                "evidence_provided": [],
                "_found_h3s": set(),  # tracks properly-styled H3 slots
            }

        def flush_tc():
            nonlocal current_tc
            if not current_tc:
                return
            if not (current_tc["test_case_id"] or current_tc["test_case_heading"]):
                return
            found_h3s = current_tc.pop("_found_h3s", set())
            missing_h3s = REQUIRED_H3S - found_h3s
            if missing_h3s:
                # One or more H3 subsections absent/unstyled → emit FAIL stub only
                test_cases.append({
                    "test_case_heading":  current_tc["test_case_heading"],
                    "test_case_id":       current_tc["test_case_id"],
                    "status":             "FAIL",
                    "missing_subsections": sorted(missing_h3s),
                })
            else:
                current_tc["test_case_name"]        = current_tc["test_case_name"].strip()
                current_tc["test_case_description"] = current_tc["test_case_description"].strip()
                # current_tc["test_observation"]      = current_tc["test_observation"].strip() # No longer a string
                current_tc.pop("_found_h3s", None)
                test_cases.append(current_tc)

        for item in section_content:
            if not isinstance(item, dict):
                continue

            item_type = item.get("type", "")
            style = item.get("style", "")
            text = (item.get("text") or "").strip()

            if item_type == "paragraph" and style == "Heading 2":
                flush_tc()
                current_tc = new_test_case(text)
                current_sub = None
                exec_order = 0
                evidence_order = 0
                continue

            if current_tc is None:
                continue

            if item_type == "paragraph" and style == "Heading 3":
                current_sub = cls._map_heading3(text)
                if current_sub:
                    current_tc["_found_h3s"].add(current_sub)
                    # Extract any trailing content from the Heading 3 paragraph itself.
                    # Try using bold_formatting as the label delimiter.
                    bold_text = (item.get("bold_formatting") or "").strip()
                    content_after = ""
                    if bold_text and text.startswith(bold_text):
                        content_after = text[len(bold_text):].strip()
                    elif ":" in text:
                        # Fallback: split at first colon if label ends with it
                        _, content_after = text.split(":", 1)
                        content_after = content_after.strip()
                    else:
                        # Aggressive fallback: use regex to find known label patterns
                        m = cls._H3_LABEL_RE.match(text)
                        if m:
                            content_after = text[m.end():].strip()

                    if content_after:
                        if current_sub == "name":
                            current_tc["test_case_name"] += content_after + " "
                        elif current_sub == "description":
                            current_tc["test_case_description"] += content_after + " "
                        elif current_sub == "execution":
                            current_tc["execution"].append({"order": exec_order, "step": content_after})
                            exec_order += 1
                        elif current_sub == "observation":
                            current_tc["test_observation"].append({"observation": content_after})
                        elif current_sub == "evidence":
                            current_tc["evidence_provided"].append({"order": evidence_order, "evidence": content_after})
                            evidence_order += 1
                continue

            if current_sub is None:
                if item_type == "paragraph" and text and not current_tc["test_case_id"]:
                    cand = cls._extract_tc_id(text)
                    if cand:
                        current_tc["test_case_id"] = cand
                continue

            if current_sub == "name":
                if item_type == "paragraph" and text:
                    current_tc["test_case_name"] += text + " "
            elif current_sub == "description":
                if item_type == "paragraph" and text:
                    current_tc["test_case_description"] += text + " "
            elif current_sub == "execution":
                if item_type == "paragraph" and text:
                    current_tc["execution"].append({"order": exec_order, "step": text})
                    exec_order += 1
                elif item_type == "image":
                    current_tc["execution"].append({"order": exec_order, "type": "image", "image_path": item.get("image_path", "")})
                    exec_order += 1
                elif item_type == "table":
                    current_tc["execution"].append({"order": exec_order, "type": "table", "rows": item.get("rows", [])})
                    exec_order += 1
            elif current_sub == "observation":
                if item_type == "paragraph" and text:
                    current_tc["test_observation"].append({"observation": text})
                elif item_type == "image":
                    current_tc["test_observation"].append({"type": "image", "image_path": item.get("image_path", "")})
                elif item_type == "table":
                    current_tc["test_observation"].append({"type": "table", "table": item.get("rows", [])})
            elif current_sub == "evidence":
                if item_type == "paragraph" and text:
                    current_tc["evidence_provided"].append({"order": evidence_order, "evidence": text})
                    evidence_order += 1
                elif item_type == "image":
                    current_tc["evidence_provided"].append({"order": evidence_order, "type": "image", "image_path": item.get("image_path", "")})
                    evidence_order += 1
                elif item_type == "table":
                    current_tc["evidence_provided"].append({"order": evidence_order, "evidence": "[Table]"})
                    evidence_order += 1

        flush_tc()
        return {"test_cases": test_cases, "total_test_cases": len(test_cases)}


# -------------------------------------------------
# ENHANCED SECTION BUILDER (SINGLE)
# -------------------------------------------------
class StructuredSectionBuilder:
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

    STRICT_ALIASES = {
        norm_name("ITSAR Section No & Name"): "1. ITSAR Section No & Name",
        norm_name("Security Requirement No & Name"): "2. Security Requirement No & Name",
        norm_name("Requirement Description"): "3. Requirement Description",
        norm_name("DUT Confirmation Details"): "4. DUT Confirmation Details",
        norm_name("DUT Details"): "4. DUT Confirmation Details",
        norm_name("DUT Configuration"): "5. DUT Configuration",
        norm_name("Preconditions"): "6. Preconditions",
        norm_name("Test Objective"): "7. Test Objective",
        norm_name("Test Plan"): "8. Test Plan",
        norm_name("Number of Test Scenarios"): "8.1. Number of Test Scenarios",
        norm_name("Test Bed Diagram"): "8.2. Test Bed Diagram",
        norm_name("Tools Required"): "8.3. Tools Required",
        norm_name("Test Execution Steps"): "8.4. Test Execution Steps",
        norm_name("Expected Results for Pass"): "9. Expected Results for Pass",
        norm_name("Expected Results"): "9. Expected Results for Pass",
        norm_name("Expected Format of Evidence"): "10. Expected Format of Evidence",
        norm_name("Test Execution"): "11. Test Execution",
        norm_name("Test Case Result"): "12. Test Case Result",
        norm_name("Test Case Results"): "12. Test Case Result",
    }

    def __init__(self, document_name: str):
        self.output = DocumentJSON(document_name)
        self.frontpage = FrontPage()
        self.current_section: Optional[Section] = None
        self.section_counter = 1
        self.frontpage_done = False
        self.in_sec11 = False

    def _next_section_id(self) -> str:
        sid = f"SEC-{self.section_counter:02d}"
        self.section_counter += 1
        return sid

    @staticmethod
    def _key_to_section_id(key: str) -> str:
        """Map a split-map key to a deterministic section_id.

        Mapping:
          front_page  → SEC-0
          sec1…sec12  → SEC-1…SEC-12
          sec8_1…sec8_4 → SEC-8-1…SEC-8-4
          tc_1, tc_2… → SEC-11-1, SEC-11-2… (resolved at call site)
        """
        if key == "front_page":
            return "SEC-0"
        import re
        m8 = re.match(r'^sec8_(\d+)$', key)
        if m8:
            return f"SEC-8-{m8.group(1)}"
        m = re.match(r'^sec(\d+)$', key)
        if m:
            return f"SEC-{m.group(1)}"
        # tc_N  →  handled in build_from_map after numbered
        mtc = re.match(r'^tc_(\d+)$', key)
        if mtc:
            return f"SEC-11-{mtc.group(1)}"
        return f"SEC-{key}"

    def _match_strict_by_name(self, text: str) -> Optional[str]:
        n = norm_name(text)
        if not n:
            return None
        if n in self.STRICT_ALIASES:
            return self.STRICT_ALIASES[n]
        for s in self.STRICT_SECTIONS:
            if norm_name(s) == n:
                return s
        return None

    def _detect_heading(self, block: Dict[str, Any]) -> Tuple[bool, Optional[int], Optional[str]]:
        if block.get("type") != "paragraph":
            return False, None, None

        style = block.get("style", "")
        raw = (block.get("text") or "").strip()
        if not raw:
            return False, None, None

        if style not in {"Heading 1", "Heading 2"}:
            return False, None, None

        canonical = self._match_strict_by_name(raw)
        level = 1 if style == "Heading 1" else 2
        title = canonical if canonical else raw
        return True, level, title

    def _get_section_type(self, title: str) -> str:
        n = norm_name(title)
        if n == norm_name("Number of Test Scenarios"):
            return "section_8_1"
        if n == norm_name("Test Execution Steps"):
            return "section_8_4"
        if n == norm_name("Expected Results for Pass") or n == norm_name("Expected Results"):
            return "section_9"
        if n == norm_name("Tools Required"):
            return "section_8_3"
        if n == norm_name("Preconditions"):
            return "section_6"
        if n == norm_name("Test Execution"):
            return "section_11"
        if n == norm_name("Test Case Result") or n == norm_name("Test Case Results"):
            return "section_12"
        if n == norm_name("DUT Configuration"):
            return "section_5"
        if n == norm_name("DUT Confirmation Details") or n == norm_name("DUT Details"):
            return "section_4"
        if n == norm_name("ITSAR Section No & Name"):
            return "section_1"
        if n == norm_name("Security Requirement No & Name"):
            return "section_2"
        if n == norm_name("Requirement Description"):
            return "section_3"
        return "standard"

    def _apply_structured_extraction(self, section: Section):
        if section.extracted:
            return

        stype = self._get_section_type(section.title)

        if stype == "section_8_1":
            section.structured_data = Section81StructuredExtractor.extract(section.content)
            section.content = []
        elif stype == "section_8_4":
            section.structured_data = Section84StructuredExtractor.extract(section.content)
            section.content = []
        elif stype == "section_9":
            section.structured_data = Section9StructuredExtractor.extract(section.content)
            section.content = []
        elif stype == "section_8_3":
            section.structured_data = Section83StructuredExtractor.extract(section.content)
            section.content = []
        elif stype == "section_6":
            section.structured_data = Section6StructuredExtractor.extract(section.content)
            section.content = []
        elif stype == "section_11":
            section.structured_data = Section11StructuredExtractor.extract(section.content)
            section.content = []
            section.extracted = True
        elif stype == "section_12":
            section.structured_data = Section12StructuredExtractor.extract(section.content)
            section.content = []
        elif stype == "section_5":
            section.structured_data = Section5StructuredExtractor.extract(section.content)
            section.content = []
        elif stype == "section_4":
            section.structured_data = Section4StructuredExtractor.extract(section.content)
            section.content = []
        elif stype == "section_1":
            section.structured_data = Section1StructuredExtractor.extract(section.content)
            section.content = []
        elif stype == "section_2":
            section.structured_data = Section2StructuredExtractor.extract(section.content)
            section.content = []
        elif stype == "section_3":
            section.structured_data = Section3StructuredExtractor.extract(section.content)
            section.content = []

    def build(self, raw_blocks: List[Dict[str, Any]]) -> DocumentJSON:
        for block in raw_blocks:
            is_heading, level, title = self._detect_heading(block)

            if is_heading:
                if not self.frontpage_done:
                    self.frontpage_done = True

                # If inside Section 11, absorb Heading 2/3 as content (for Section11StructuredExtractor)
                # Only break out when we hit a new Heading 1 (e.g. Section 12)
                if self.in_sec11:
                    if level == 1:
                        # Leaving Section 11 — close it and proceed
                        if self.current_section:
                            self._apply_structured_extraction(self.current_section)
                        self.in_sec11 = False
                        # Fall through to create new section below
                    else:
                        # Heading 2 or 3 inside section 11 — add as content
                        if self.current_section is not None:
                            self.current_section.content.append({
                                "type": "paragraph",
                                "text": block.get("text", ""),
                                "style": block.get("style", ""),
                                "bold_formatting": block.get("bold_formatting"),
                            })
                        continue

                if not self.in_sec11:
                    if self.current_section:
                        if self._get_section_type(self.current_section.title) != "standard":
                            self._apply_structured_extraction(self.current_section)

                section = Section(section_id=self._next_section_id(), title=title, level=level)
                self.output.sections.append(section)
                self.current_section = section

                # Track if we just entered Section 11
                if self._get_section_type(title) == "section_11":
                    self.in_sec11 = True

                continue

            if not self.frontpage_done:
                if block.get("type") == "paragraph" and (block.get("text") or "").strip():
                    self.frontpage.content.append(block.get("text"))
                continue

            if self.current_section is None:
                continue

            btype = block.get("type")

            if btype == "paragraph":
                txt = block.get("text") or ""
                sty = block.get("style") or ""
                if txt.strip() or sty.startswith("Heading"):
                    self.current_section.content.append({
                        "type": "paragraph",
                        "text": txt,
                        "style": sty,
                        "bold_formatting": block.get("bold_formatting"),
                    })

            elif btype == "image":
                self.current_section.content.append({
                    "type": "image",
                    "image_path": block.get("path", ""),
                })

            elif btype == "table":
                self.current_section.content.append({
                    "type": "table",
                    "rows": block.get("rows", []),
                    "style": block.get("style", ""),
                })

        if self.current_section:
            if self._get_section_type(self.current_section.title) != "standard":
                self._apply_structured_extraction(self.current_section)

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
        Index [0] is always the heading block; content starts at index [1].

        This method:
          - Does NOT re-scan the full document.
          - Does NOT re-detect headings.
          - Calls each section's typed extractor on its exact block slice only.
          - Completely eliminates section-bleed between adjacent sections.
        """
        # Section key → canonical title + level
        SEC_META: Dict[str, tuple] = {
            "front_page": ("[Front Page]", 0),
            "sec1":  ("1. ITSAR Section No & Name",           1),
            "sec2":  ("2. Security Requirement No & Name",    1),
            "sec3":  ("3. Requirement Description",           1),
            "sec4":  ("4. DUT Confirmation Details",          1),
            "sec5":  ("5. DUT Configuration",                 1),
            "sec6":  ("6. Preconditions",                     1),
            "sec7":  ("7. Test Objective",                    1),
            "sec8":  ("8. Test Plan",                         1),
            "sec8_1": ("8.1. Number of Test Scenarios",       2),
            "sec8_2": ("8.2. Test Bed Diagram",               2),
            "sec8_3": ("8.3. Tools Required",                 2),
            "sec8_4": ("8.4. Test Execution Steps",           2),
            "sec9":  ("9. Expected Results for Pass",         1),
            "sec10": ("10. Expected Format of Evidence",      1),
            "sec11": ("11. Test Execution",                   1),
            "sec12": ("12. Test Case Result",                 1),
        }

        # Helper: convert a raw block slice (starting at idx 1) into content dicts
        def _to_content(raw_slice: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
            content = []
            for b in raw_slice[1:]:          # skip heading block
                btype = b.get("type", "")
                if btype == "paragraph":
                    content.append({
                        "type": "paragraph",
                        "text": b.get("text", ""),
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

        # Ordered rendering: sec1 … sec12, each with possible sub-sections
        ORDERED_KEYS = [
            "sec1", "sec2", "sec3", "sec4", "sec5", "sec6", "sec7",
            "sec8", "sec8_1", "sec8_2", "sec8_3", "sec8_4",
            "sec9", "sec10", "sec11", "sec12",
        ]

        # Collect tc_* keys for Section 11 test cases, sorted
        tc_keys = sorted(
            [k for k in mapped_sections if k.startswith("tc_")],
            key=lambda k: int(k.split("_")[1]),
        )

        # Track which sec11 tc blocks have already been rendered inside sec11
        rendered_tc = set()

        failed_keys = failed_keys or set()

        # Check if Section 9 is missing in Header 1
        sec9_missing = ("sec9" in failed_keys) or ("sec9" not in mapped_sections)

        for key in ORDERED_KEYS:
            raw_slice = mapped_sections.get(key)
            if raw_slice is None:
                # Emit a FAIL stub if this key was explicitly failed
                if key in failed_keys and key in SEC_META:
                    canonical_title = SEC_META[key][0]
                    level = SEC_META[key][1]
                    stub = Section(
                        section_id=self._key_to_section_id(key),
                        title=canonical_title,
                        level=level,
                    )
                    stub.structured_data = {"status": "FAIL"}
                    self.output.sections.append(stub)
                continue

            meta = SEC_META.get(key, (key, 1))
            canonical_title = meta[0]
            level = meta[1]

            # Get content (skip heading at [0])
            content = _to_content(raw_slice)

            section = Section(
                section_id=self._key_to_section_id(key),
                title=canonical_title,
                level=level,
            )

            stype = self._get_section_type(canonical_title)

            # GUARD: If Section 9 is missing in Header 1, Section 8 sub-sections must also fail 
            # to avoid content bleed (since Section 9 content would be inside 8.4).
            if key.startswith("sec8"):
                if sec9_missing:
                    section.structured_data = {"status": "FAIL"}
                    self.output.sections.append(section)
                    continue

            if stype == "section_8_1":
                section.structured_data = Section81StructuredExtractor.extract(content)
            elif stype == "section_8_4":
                section.structured_data = Section84StructuredExtractor.extract(content)
            elif stype == "section_9":
                section.structured_data = Section9StructuredExtractor.extract(content)
            elif stype == "section_8_3":
                section.structured_data = Section83StructuredExtractor.extract(content)
            elif stype == "section_6":
                section.structured_data = Section6StructuredExtractor.extract(content)
            elif stype == "section_12":
                section.structured_data = Section12StructuredExtractor.extract(content)
            elif stype == "section_5":
                section.structured_data = Section5StructuredExtractor.extract(content)
            elif stype == "section_4":
                section.structured_data = Section4StructuredExtractor.extract(content)
            elif stype == "section_1":
                section.structured_data = Section1StructuredExtractor.extract(content)
            elif stype == "section_2":
                section.structured_data = Section2StructuredExtractor.extract(content)
            elif stype == "section_3":
                section.structured_data = Section3StructuredExtractor.extract(content)
            elif stype == "section_11":
                # Section 11: assemble content only from tc_* slices that PASSED.
                # failed tc_* keys were removed from mapped_sections by extract_document.py.
                _failed_tc_hdgs = failed_tc_headings or {}

                # Full ordered key list: passing (from mapped_sections) + failing, sorted by number.
                _all_tc_keys = sorted(
                    list(tc_keys) + list(_failed_tc_hdgs.keys()),
                    key=lambda k: int(k.split("_")[1]) if k.split("_")[1].isdigit() else 99999,
                )

                # Collect content blocks only from passing TCs for the extractor.
                sec11_content: List[Dict[str, Any]] = []
                sec11_content.extend(content)
                for tc_key in tc_keys:                  # tc_keys = only passing keys
                    tc_slice = mapped_sections.get(tc_key, [])
                    for b in tc_slice:
                        btype = b.get("type", "")
                        if btype == "paragraph":
                            sec11_content.append({
                                "type": "paragraph",
                                "text": b.get("text", ""),
                                "style": b.get("style", ""),
                                "bold_formatting": b.get("bold_formatting"),
                            })
                        elif btype == "table":
                            sec11_content.append({"type": "table", "rows": b.get("rows", [])})
                        elif btype == "image":
                            sec11_content.append({"type": "image", "image_path": b.get("path", "")})
                    rendered_tc.add(tc_key)

                raw_extracted = Section11StructuredExtractor.extract(sec11_content)

                # Stamp passing TCs with their SEC-11-N ids (positional, so order matters).
                passing_tc_list = raw_extracted.get("test_cases", [])
                for tc_data, tc_key in zip(passing_tc_list, tc_keys):
                    tc_num = tc_key.split("_")[-1]
                    tc_data["section_id"] = f"SEC-11-{tc_num}"

                # Build a fast lookup: section_id -> tc_data (for PASS TCs)
                _pass_by_sid: Dict[str, Any] = {
                    td["section_id"]: td for td in passing_tc_list if "section_id" in td
                }

                # Assemble the final ordered list: tc_1, tc_2, tc_3 ... interleaved.
                ordered_test_cases: List[Dict[str, Any]] = []
                for tc_key in _all_tc_keys:
                    tc_num = tc_key.split("_")[-1]
                    sid = f"SEC-11-{tc_num}"
                    if tc_key in tc_keys:
                        # PASS: use the extracted data (already stamped with section_id)
                        tc_data = _pass_by_sid.get(sid)
                        if tc_data:
                            ordered_test_cases.append(tc_data)
                    else:
                        # FAIL: emit a minimal stub.
                        # Title = exact H2 text from the document.
                        actual_heading = _failed_tc_hdgs.get(tc_key, tc_key)
                        ordered_test_cases.append({
                            "section_id": sid,
                            "test_case_heading": actual_heading,
                            "test_case_id": "",
                            "status": "FAIL",
                        })

                section.structured_data = {
                    "test_cases": ordered_test_cases,
                    "total_test_cases": len(ordered_test_cases),
                }
                section.extracted = True
            else:
                # Generic: keep as content list
                section.content = content

            self.output.sections.append(section)

        return self.output


# -------------------------------------------------
# MAIN ENTRY POINT
# -------------------------------------------------
def build_structured_document(lossless_json_path: Path) -> Dict[str, Any]:
    with open(lossless_json_path, "r", encoding="utf-8") as f:
        lossless_data = json.load(f)

    document_name = lossless_data["document"]
    raw_blocks = lossless_data["blocks"]

    builder = StructuredSectionBuilder(document_name)
    doc = builder.build(raw_blocks)
    return doc.to_dict()


def main():
    if len(sys.argv) < 2:
        print("Usage: python AI_structured_extract.py <lossless.json>")
        sys.exit(1)

    lossless_path = Path(sys.argv[1])
    if not lossless_path.exists():
        print(f"Error: File not found: {lossless_path}")
        sys.exit(1)

    structured_doc = build_structured_document(lossless_path)
    out_path = lossless_path.parent / f"{lossless_path.stem.replace('_lossless', '')}ai_structured.json"

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(structured_doc, f, indent=2, ensure_ascii=False)

    print(f"✅ AI STRUCTURED DOCUMENT CREATED: {out_path}")
    print(f"📊 Total sections: {len(structured_doc.get('sections', []))}")


if __name__ == "__main__":
    main()