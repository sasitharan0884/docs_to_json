"""
Document Extractor - Complete Pipeline
Extracts DOCX to both lossless and structured JSON in one run.
"""
import json
import sys
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from dataclasses import dataclass
from lossless_extract import parse_docx
from structured_extract import StructuredSectionBuilder, validate_frontpage_headings
from AI_structured_extract import StructuredSectionBuilder as AIStructuredSectionBuilder
from document_split_validator import run_validator


@dataclass
class ValidationIssue:
    where: str
    what: str
    suggestion: str
    severity: str = "High"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "where": self.where,
            "what": self.what,
            "suggestion": self.suggestion,
            "severity": self.severity,
        }


class ValidationError(Exception):
    def __init__(self, issues: List[ValidationIssue]):
        self.issues = issues
        super().__init__("Validation failed")


def norm_alnum(text: str) -> str:
    return re.sub(r"[^a-zA-Z0-9]+", "", (text or "")).lower()

def norm_alpha(text: str) -> str:
    return re.sub(r"[^a-zA-Z]+", "", (text or "")).lower()


def is_section11_test_case_heading2(title: str) -> bool:
    """True if Section 11 Heading 2 represents 'Test Case Number'."""
    return norm_alpha(title).startswith("testcasenumber")

def is_top_level_numbered(title: str) -> bool:
    return bool(re.match(r"^\s*\d+\.\s", title or "")) and not bool(re.match(r"^\s*\d+\.\d+\.", title or ""))

def get_heading_blocks(blocks: List[Dict[str, Any]], style: str) -> List[Tuple[int, str]]:
    return [(i, (b.get("text", "") or "")) for i, b in enumerate(blocks) if b.get("style") == style]

def find_text_occurrences(blocks: List[Dict[str, Any]], needle_norm: str) -> List[Tuple[int, str, str, str]]:
    """
    Returns hits as: (index, paragraph_style, rendered_style, text)
    We use this to detect cases where heading is applied as character style
    (rendered_style like 'Heading 3 Char') but paragraph style is not Heading X.
    """
    hits: List[Tuple[int, str, str, str]] = []
    for i, b in enumerate(blocks):
        txt = (b.get("text", "") or "").strip()
        if not txt:
            continue

        if norm_alpha(txt) != needle_norm:
            continue

        p_style = (b.get("style", "") or "").strip()
        r_style = (b.get("rendered_style", "") or "").strip()
        hits.append((i, p_style, r_style, txt))
    return hits


SCENARIO_HEADER_RE = re.compile(
    r"^\s*test\s*(scenario|case)s?\s+\d+(?:\.\d+){2,}\b",
    re.IGNORECASE
)

# Detect malformed scenario headers where there is no space before the id,
# e.g. "TestScenario1.1.1.1" or "Test Case1.1.1.1".
SCENARIO_HEADER_NO_SPACE_RE = re.compile(
    r"^\s*test\s*(scenario|case)s?\d+(?:\.\d+){2,}\b",
    re.IGNORECASE
)


def collect_scenario_headers_in_range(
    blocks: List[Dict[str, Any]], start: int, end: int
) -> Tuple[List[str], List[str]]:
    """
    Collect scenario headers and malformed scenario-like headers:
    - valid_headers: bold paragraph text matching SCENARIO_HEADER_RE
    - malformed_no_space_headers: bold paragraph text matching
      SCENARIO_HEADER_NO_SPACE_RE (missing space before id)
    Any other bold text is allowed but ignored.
    """
    valid_headers: List[str] = []
    malformed_no_space_headers: List[str] = []
    for b in blocks[start:end]:
        if b.get("type") != "paragraph":
            continue
        bold = (b.get("bold_formatting") or "").strip()
        if not bold:
            continue
        text = (b.get("text") or "").strip()
        if SCENARIO_HEADER_RE.match(text):
            valid_headers.append(bold)
        elif SCENARIO_HEADER_NO_SPACE_RE.match(text):
            malformed_no_space_headers.append(text)
    return valid_headers, malformed_no_space_headers


def extract_expected_heading1_titles() -> List[str]:
    # Only top-level (1., 2., ...), from your template
    from structured_extract import StructuredSectionBuilder
    return [s for s in StructuredSectionBuilder.STRICT_SECTIONS if is_top_level_numbered(s)]


def locate_section8_and_9(h1_blocks: List[Tuple[int, str]]) -> Tuple[Optional[int], Optional[int]]:
    """
    Returns (section8_index_in_blocks, section9_index_in_blocks) based on Heading 1 text.
    Uses name-based matching: "Test Plan" for 8, and "Expected Results for Pass" for 9.
    """
    sec8_idx = None
    sec9_idx = None

    # Matching by name ONLY (ignoring numbers)
    for idx, txt in h1_blocks:
        norm_txt = norm_alpha(txt)
        if norm_txt == "testplan":
            sec8_idx = idx
        if norm_txt in ("expectedresultsforpass", "expectedresults"):
            sec9_idx = idx

    return sec8_idx, sec9_idx


def locate_section11_and_12(h1_blocks: List[Tuple[int, str]]) -> Tuple[Optional[int], Optional[int]]:
    """
    Returns (section11_block_index, section12_block_index) based on Heading 1 text.
    Sections are matched strictly by name.
    """
    sec11_idx = None
    sec12_idx = None

    for idx, txt in h1_blocks:
        norm_txt = norm_alpha(txt)
        if norm_txt == "testexecution":
            sec11_idx = idx
        if norm_txt in ("testcaseresult", "testcaseresults"):
            sec12_idx = idx

    return sec11_idx, sec12_idx
# ---------------------------------------------------------------------------
# Section key → section number helper (sec1→1, sec2→2, ..., sec12→12)
# ---------------------------------------------------------------------------
_SEC_KEY_TO_NUM: Dict[str, int] = {f"sec{i}": i for i in range(1, 13)}




# ---------------------------------------------------------------------------
_SEC_KEY_TO_NUM: Dict[str, int] = {f"sec{i}": i for i in range(1, 13)}


def _get_failed_section_keys(validator_result: Dict[str, Any]) -> List[str]:
    """
    Return a list of section keys (e.g. ['sec2', 'sec5', 'tc_1']) whose Phase-1
    validation status is FAIL (extraction was BLOCKED).

    A section is considered failed when:
      • validator_result["sections"][key]["validation"] == "FAIL"   OR
      • validator_result["sections"][key]["extraction"] == "BLOCKED"

    WARN sections are treated as passing — they still get extracted.
    """
    failed: List[str] = []
    
    # 1. Top-level sections and Section 8 subsections
    sections = validator_result.get("sections") or {}
    for sec_key, sec_info in sections.items():
        validation = (sec_info.get("validation") or "").upper()
        extraction = (sec_info.get("extraction") or "").upper()
        if validation == "FAIL" or extraction in ("BLOCKED", "SKIPPED"):
            failed.append(sec_key)
            
        # Check sub-sections (like sec8_1, sec8_2)
        subsections = sec_info.get("subsections") or {}
        for sub_key, sub_info in subsections.items():
            sub_val = (sub_info.get("validation") or "").upper()
            sub_ext = (sub_info.get("extraction") or "").upper()
            if sub_val == "FAIL" or sub_ext == "BLOCKED":
                failed.append(sub_key)
                
    # 2. Test Cases from Phase 1 — individual tc_* keys.
    #    A TC is blocked if ANY H3 has status == "FAIL" (missing OR unstyled).
    #    This enforces the Initial Check: a TC is only created when it fully
    #    passes — all 5 H3 subsections must be present AND correctly styled.
    phase1 = validator_result.get("phase1") or {}
    p1_sections = phase1.get("sections") or {}
    sec11_info = p1_sections.get("sec11") or {}
    h2_test_cases = sec11_info.get("heading_tree", {}).get("h2_test_cases", {})

    failed_tc_keys: List[str] = []
    for tc_key, tc_info in h2_test_cases.items():
        h3_children = tc_info.get("h3_children", {})
        has_high_issue = False

        # Fail the TC if its overall Phase-1 status is already FAIL.
        # This covers: unstyled H2 heading, orphan TC, empty TC, etc.
        if tc_info.get("status") == "FAIL":
            has_high_issue = True

        # Fail the TC if the H2 heading itself was not found.
        if tc_info.get("found") is None and not tc_info.get("orphan", False):
            has_high_issue = True

        # Block extraction if ANY H3 has status == "FAIL".
        # This covers both completely absent H3s (found is None) AND
        # H3s that were present but lacked the correct Heading 3 style.
        has_failing_h3 = any(
            v.get("status") == "FAIL"
            for v in h3_children.values()
        )
        if has_failing_h3:
            has_high_issue = True

        if has_high_issue:
            failed_tc_keys.append(tc_key)

    failed.extend(failed_tc_keys)

    # 3. If ALL test cases in sec11 failed, also exclude sec11 itself.
    #    Without this, the sec11 heading block (NOT inside any tc_* range)
    #    would survive filtering and produce an empty "11. Test Execution"
    #    section in structured / ai_structured output.
    if h2_test_cases and len(failed_tc_keys) == len(h2_test_cases):
        if "sec11" not in failed:
            failed.append("sec11")

    return failed

_ORDINAL_NAMES = [
    "zero", "first", "second", "third", "fourth", "fifth",
    "sixth", "seventh", "eighth", "ninth", "tenth",
    "eleventh", "twelfth",
]

_SEC_KEYS_ORDERED = [
    "sec1", "sec2", "sec3", "sec4", "sec5", "sec6",
    "sec7", "sec8", "sec9", "sec10", "sec11", "sec12",
]

_SEC8_H2_KEYS_ORDERED = ["sec8_1", "sec8_2", "sec8_3", "sec8_4"]

_H3_ALPHA_KEYS = ["a", "b", "c", "d", "e"]


def _split_block_ids(range_str: str) -> List[int]:
    """Parse '25-89' → [25, 26, … 89]. Returns [] if unparseable."""
    if not range_str or "-" not in range_str:
        return []
    try:
        start, end = map(int, range_str.split("-", 1))
        return list(range(start, end + 1))
    except ValueError:
        return []



def _h3_splitter_for_test_case(
    h3_tree: Dict[str, Any],
) -> Dict[str, Any]:
    """
    Build h3_splitter for one test case.

    Returns:
        {
          "a. Test Case Name": {"status": "PASS"|"FAIL", "found": "..."|None},
          "b. Test Case Description": {…},
          …
        }
    """
    result: Dict[str, Any] = {}
    for key in _H3_ALPHA_KEYS:
        h3_info = h3_tree.get(key, {})
        expected_name = h3_info.get("expected", key)
        found_heading = h3_info.get("found")
        status = "PASS" if h3_info.get("status") == "PASS" else "FAIL"
        result[f"{key}. {expected_name}"] = {
            "status": status,
            "found": found_heading,
        }
    return result


def build_checklist_output(
    validator_result: Dict[str, Any],
    lossless_data: Dict[str, Any],
) -> Dict[str, Any]:
    """
    Build a hierarchical checklist output from the validator result.

    Structure
    ---------
    {
      "zero_h1":   { "block_ids": [0…43], "status": "PASS", "heading": "[Front Page]" },
      "first_h1":  { "block_ids": [44…46], "status": "PASS", "heading": "1. ITSAR …",
                     "section_key": "sec1",
                     "validation": "PASS", "issues": [] },
      …
      "eighth_h1": { …,
        "eight_h2validator": {
          "sec8_1": { "heading": "8.1 Number of Test Scenarios", "block_ids": […],
                      "status": "PASS", "test_scenarios": […] },
          "sec8_2": { … },
          "sec8_3": { … },
          "sec8_4": { … }
        }
      },
      …
      "eleventh_h1": { …,
        "h2_splitter": {
          "Test Case Number 1.1.4.1": {
            "block_ids": […], "status": "PASS",
            "h3_splitter": {
              "a. Test Case Name": { "status": "PASS", "found": "a. Test Case Name",
                "h4_splitter": { "CLI": […], "content": […] } },
              …
            }
          },
          …
        }
      },
      "twelfth_h1": { … },

      "zero_sec_validator": {
        "description": "runner1_apicall({}, {}, project_details) — overall pipeline result",
        "status": "PASS"|"PARTIAL"|"FAIL",
        "total_sections": 12,
        "extracted": N,
        "blocked_invalid": N,
        "extraction_errors": N,
        "section_summary": { "sec1": "PASS", "sec2": "WARN", … }
      }
    }
    """
    blocks: List[Dict[str, Any]] = lossless_data.get("blocks") or []
    split_map_detail: Dict[str, Any] = validator_result.get("split_map_detail") or {}
    phase1_sections: Dict[str, Any] = validator_result.get("phase1", {}).get("sections") or {}
    unified_sections: Dict[str, Any] = validator_result.get("sections") or {}
    summary: Dict[str, Any] = validator_result.get("summary") or {}

    # ---------------------------------------------------------------
    # zero_h1  — front page (before first H1)
    # ---------------------------------------------------------------
    fp_detail = split_map_detail.get("front_page") or {}
    fp_range = fp_detail.get("range", "")
    zero_h1: Dict[str, Any] = {
        "heading": fp_detail.get("heading", "[Front Page]"),
        "block_ids": _split_block_ids(fp_range),
        "block_range": fp_range,
        "status": "PASS",
    }

    checklist: Dict[str, Any] = {"zero_h1": zero_h1}

    # ---------------------------------------------------------------
    # first_h1 … twelfth_h1  — each of the 12 H1 sections
    # ---------------------------------------------------------------
    for sec_idx, sec_key in enumerate(_SEC_KEYS_ORDERED, start=1):
        ordinal = _ORDINAL_NAMES[sec_idx]   # "first" … "twelfth"
        var_name = f"{ordinal}_h1"

        p1_sec = phase1_sections.get(sec_key) or {}
        detail = split_map_detail.get(sec_key) or {}
        unified = unified_sections.get(sec_key) or {}

        range_str = detail.get("range", p1_sec.get("block_range", "N/A"))
        heading = detail.get("heading") or p1_sec.get("found") or p1_sec.get("expected", sec_key)
        validation = p1_sec.get("status", "FAIL")
        issues = p1_sec.get("issues") or []

        sec_node: Dict[str, Any] = {
            "section_key": sec_key,
            "heading": heading,
            "block_range": range_str,
            "block_ids": _split_block_ids(range_str),
            "status": validation,
            "issues": issues,
        }

        # ---- Section 8: eight_h2validator ----
        if sec_key == "sec8":
            h2_validator: Dict[str, Any] = {}
            h2_tree = (p1_sec.get("heading_tree") or {}).get("h2_children") or {}

            for sub_key in _SEC8_H2_KEYS_ORDERED:
                sub_detail = split_map_detail.get(sub_key) or {}
                sub_tree = h2_tree.get(sub_key) or {}
                sub_range = sub_detail.get("range", sub_tree.get("block_range", "N/A"))
                sub_heading = sub_detail.get("heading") or sub_tree.get("found") or sub_tree.get("expected", sub_key)
                sub_status = sub_tree.get("status", "FAIL")
                sub_scenarios = sub_tree.get("test_scenarios") or []

                h2_validator[sub_key] = {
                    "heading": sub_heading,
                    "block_range": sub_range,
                    "block_ids": _split_block_ids(sub_range),
                    "status": sub_status,
                    "num_prefix_ok": sub_tree.get("num_prefix", False),
                    "test_scenarios": sub_scenarios,
                }

            sec_node["eight_h2validator"] = h2_validator

        # ---- Section 11: h2_splitter → h3_splitter ----
        if sec_key == "sec11":
            h2_test_cases = (p1_sec.get("heading_tree") or {}).get("h2_test_cases") or {}
            h2_splitter: Dict[str, Any] = {}

            tc_detail_map: Dict[str, Any] = {
                k: v for k, v in split_map_detail.items() if k.startswith("tc_")
            }

            for tc_key in sorted(tc_detail_map.keys(), key=lambda k: int(k.split("_")[1])):
                tc_info = h2_test_cases.get(tc_key) or {}
                tc_detail = tc_detail_map[tc_key]
                tc_range = tc_detail.get("range", "N/A")
                tc_heading = tc_detail.get("heading") or tc_info.get("heading", tc_key)
                tc_status = tc_info.get("status", "FAIL")
                tc_h3_tree = tc_info.get("h3_children") or {}

                tc_block_ids = _split_block_ids(tc_range)

                # h3_splitter for this test case
                h3_split = _h3_splitter_for_test_case(tc_h3_tree)

                h2_splitter[tc_heading] = {
                    "tc_key": tc_key,
                    "block_range": tc_range,
                    "block_ids": tc_block_ids,
                    "status": tc_status,
                    "h3_splitter": h3_split,
                }

            sec_node["h2_splitter"] = h2_splitter

        checklist[var_name] = sec_node

    # ---------------------------------------------------------------
    # zero_sec_validator — wraps runner1_apicall({}, {}, project_details)
    # ---------------------------------------------------------------
    section_summary: Dict[str, str] = {}
    for sec_key in _SEC_KEYS_ORDERED:
        p1 = phase1_sections.get(sec_key) or {}
        section_summary[sec_key] = p1.get("status", "MISSING")

    checklist["zero_sec_validator"] = {
        "description": "runner1_apicall({}, {}, project_details) — overall pipeline validation",
        "status": summary.get("status", "FAIL"),
        "total_sections": summary.get("total_sections", 12),
        "extracted": summary.get("extracted", 0),
        "blocked_invalid": summary.get("blocked_invalid", 0),
        "extraction_errors": summary.get("extraction_errors", 0),
        "section_summary": section_summary,
    }

    return checklist


# ---------------------------------------------------------------------------
# MAIN PIPELINE
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_document.py <input.docx>")
        sys.exit(1)

    file_path = Path(sys.argv[1])
    if not file_path.exists():
        print(f"Error: File not found: {file_path}")
        sys.exit(1)

    base_name  = file_path.stem
    output_dir = file_path.parent

    # ==================================================================
    # STEP 1 — Lossless extraction  (always runs, always saved)
    # ==================================================================
    print(f"\n{'='*60}")
    print(f"STEP 1 — Lossless extraction")
    print(f"{'='*60}")
    try:
        lossless_data = parse_docx(file_path, mode="lossless")
    except Exception as exc:
        print(f"[FATAL] Lossless extraction failed: {exc}")
        sys.exit(1)

    lossless_path = output_dir / f"{base_name}_lossless.json"
    with open(lossless_path, "w", encoding="utf-8") as f:
        json.dump(lossless_data, f, indent=2, ensure_ascii=False)
    print(f"  Saved: {lossless_path}")
    print(f"  Total blocks: {lossless_data.get('total_blocks', len(lossless_data.get('blocks', [])))}")

    # ==================================================================
    # STEP 2 — document_split_validator  (always runs, always saved)
    #   • Validates every section (1-12) independently
    #   • PASS / WARN → section extracted in output.json
    #   • FAIL        → section blocked in output.json
    # ==================================================================
    print(f"\n{'='*60}")
    print(f"STEP 2 — document_split_validator")
    print(f"{'='*60}")
    try:
        validator_result = run_validator(lossless_data)
    except Exception as exc:
        print(f"[FATAL] Validator crashed unexpectedly: {exc}")
        sys.exit(1)

    debug_json_path = output_dir / "debug.json"
    with open(debug_json_path, "w", encoding="utf-8") as f:
        json.dump(validator_result, f, indent=2, ensure_ascii=False)
    print(f"  Saved: {debug_json_path}")

    # ==================================================================
    # STEP 2b — Build Structured Validation Summary (output.json)
    # ==================================================================
    _EXPECTED_SECTIONS = [
        ("sec1", "1. ITSAR Section No & Name"),
        ("sec2", "2. Security Requirement No & Name"),
        ("sec3", "3. Requirement Description"),
        ("sec4", "4. DUT Confirmation Details"),
        ("sec5", "5. DUT Configuration"),
        ("sec6", "6. Preconditions"),
        ("sec7", "7. Test Objective"),
        ("sec8", "8. Test Plan"),
        ("sec9", "9. Expected Results for Pass"),
        ("sec10", "10. Expected Format of Evidence"),
        ("sec11", "11. Test Execution"),
        ("sec12", "12. Test Case Result"),
    ]

    sections_result = validator_result.get("sections", {})
    # Get failed keys early to use them for skipping
    failed_keys = _get_failed_section_keys(validator_result)

    structured_sections = []
    all_format_checks = {}

    phase1_sections = validator_result.get("phase1", {}).get("sections", {})

    # ------------------------------------------------------------------
    # section_0 — Front Page (blocks before the first H1)
    # If section_1 (sec1) is FAIL/MISSING, the front page end-boundary
    # is unknown, so section_0 also cascades to FAIL.
    # ------------------------------------------------------------------
    sec1_p1   = phase1_sections.get("sec1", {})
    sec1_status = sec1_p1.get("status", "PASS")
    sec1_failed = sec1_status in ("FAIL", "MISSING")

    fp_detail    = validator_result.get("split_map_detail", {}).get("front_page") or {}
    fp_range     = fp_detail.get("range", "N/A")

    # --- Front Page Heading 2 validation -----------------------------------
    # Run independently of the sec1 boundary check so both error types can
    # appear together in section_0 when both conditions are triggered.
    # NOTE: the local `mapped_sections` is built later; read from
    # validator_result directly to avoid a forward-reference error.
    _fp_raw_blocks = (validator_result.get("mapped_sections") or {}).get("front_page", [])
    _fp_h2_result  = validate_frontpage_headings(_fp_raw_blocks)
    _fp_h2_failed  = _fp_h2_result["status"] == "FAIL"

    section0_errors: list = []

    if sec1_failed:
        section0_errors.append({
            "type":          "BOUNDARY_UNKNOWN",
            "severity":      "HIGH",
            "message":       (
                "Front Page end-boundary cannot be determined because "
                "Section 1 ('1. ITSAR Section No & Name') is missing. "
                "The Front Page content may be contaminated or incomplete."
            ),
            "suggestion":    "Add a Heading 1 titled '1. ITSAR Section No & Name' so the Front Page boundary can be reliably identified.",
            "where":         "Front Page",
            "redirect_text": "Front Page",
            "what":          "Missing section boundary caused by absent Section 1 heading",
        })

    if _fp_h2_failed:
        missing_labels = _fp_h2_result.get("missing", [])
        section0_errors.append({
            "type":          "MISSING_FRONTPAGE_HEADINGS",
            "severity":      "HIGH",
            "message":       (
                f"Front Page is missing {len(missing_labels)} required Heading 2 "
                f"field(s): {', '.join(missing_labels)}. "
                "All 7 DUT metadata fields must be present as Heading 2 paragraphs "
                "in the correct order."
            ),
            "suggestion":    (
                "Ensure the following Heading 2 labels are present in the Front Page in order: "
                + ", ".join([
                    "DUT Details:", "DUT Software Version:", "Digest Hash of OS:",
                    "Digest Hash of Configuration:", "Applicable ITSAR:",
                    "ITSAR Version No:", "OEM Supplied Document list:",
                ])
            ),
            "where":         "Front Page",
            "redirect_text": "Front Page",
            "what":          f"Missing required Heading 2 field(s): {', '.join(missing_labels)}",
            "missing_fields": missing_labels,
            "found_order":    _fp_h2_result.get("found_order", []),
        })

    section0_status   = "FAIL" if section0_errors else "PASS"
    section0_findings = "Issues found." if section0_errors else "No findings."

    section0_obj = {
        "section_id":   "section_0",
        "section_name": "Front Page",
        "checks": [
            {
                "check_name": "Heading",
                "validation_results": [
                    {
                        "checklist_name": "Section Structure & Completeness",
                        "status":         section0_status,
                        "error_count":    len(section0_errors),
                        "errors":         section0_errors,
                        "findings":       section0_findings,
                    }
                ],
            }
        ],
    }
    structured_sections.append(section0_obj)
    # ------------------------------------------------------------------

    for i, (sec_key, sec_name) in enumerate(_EXPECTED_SECTIONS):
        sec_info = sections_result.get(sec_key, {})
        issues = sec_info.get("issues", [])
        
        p1_sec = phase1_sections.get(sec_key, {})
        heading_tree = p1_sec.get("heading_tree", {})
        h2_dict = heading_tree.get("h2_children") or heading_tree.get("h2_test_cases") or {}

        # Separate issues into "general" vs "subsection"
        general_errors = []
        sub_errors = {k: [] for k in h2_dict.keys()}

        for iss in issues:
            assigned = False
            for sub_key, sub_info in h2_dict.items():
                sub_name = sub_info.get("found") or sub_info.get("expected") or ""
                # Check if the issue message mentions the subsection name
                if sub_name and sub_name in iss.get("message", ""):
                    iss_copy = iss.copy()
                    iss_copy["where"] = sub_name
                    iss_copy["redirect_text"] = sub_name
                    if "what" not in iss_copy:
                        iss_copy["what"] = "Missing or invalid section formatting"
                    sub_errors[sub_key].append(iss_copy)
                    assigned = True
                    break
            
            if not assigned:
                iss_copy = iss.copy()
                iss_copy["where"] = iss.get("where", sec_name)
                iss_copy["redirect_text"] = iss.get("where", sec_name)
                if "what" not in iss_copy:
                    iss_copy["what"] = "Missing or invalid section formatting"
                general_errors.append(iss_copy)

        val_results = []

        gen_status = "FAIL" if general_errors else "PASS"
        findings_text = "Issues found." if general_errors else "No findings."

        val_results.append({
            "checklist_name": "Section Structure & Completeness",
            "status": gen_status,
            "error_count": len(general_errors),
            "errors": general_errors,
            "findings": findings_text
        })

        # For sec8: H2 subsections (8.1–8.4) become nested inside section_8.
        # For sec11: test cases become nested inside section_11.
        if sec_key in ("sec8", "sec11"):
            # Prepare parent section object
            sec_obj = {
                "section_id": f"section_{i+1}",
                "section_name": sec_name,
                "checks": [{"check_name": "Heading", "validation_results": val_results}],
                "subsections": [] # Nest children here
            }
            
            SEC8_EXPECTED = [
                ("sec8_1", "8.1. Number of Test Scenarios", "section_8_1"),
                ("sec8_2", "8.2. Test Bed Diagram", "section_8_2"),
                ("sec8_3", "8.3. Tools Required", "section_8_3"),
                ("sec8_4", "8.4. Test Execution Steps", "section_8_4"),
            ]

            # Use hardcoded list for Section 8, but dynamic h2_dict for Section 11
            if sec_key == "sec8":
                items_to_process = []
                for sub_key, sub_name, sub_id in SEC8_EXPECTED:
                    sub_info = h2_dict.get(sub_key, {})
                    items_to_process.append((sub_key, sub_name, sub_id, sub_info))
            else:
                # Dynamic for Section 11
                items_to_process = []
                for sub_key, sub_info in h2_dict.items():
                    # h2_test_cases entries store the actual heading text under "found"
                    sub_name = sub_info.get("found") or sub_info.get("heading") or sub_info.get("expected") or sub_key
                    sub_id = f"section_11_{sub_key.split('_')[-1]}" if sub_key.startswith("tc_") else f"section_{sub_key}"
                    items_to_process.append((sub_key, sub_name, sub_id, sub_info))

            for sub_key, sub_name, sub_sec_id, sub_info in items_to_process:
                sub_errs = sub_errors.get(sub_key, [])
                sub_status = "FAIL" if sub_errs else "PASS"
                sub_findings = "Issues found." if sub_errs else "No findings."

                sub_val_results = [{
                    "checklist_name": "Section Structure & Completeness",
                    "status": sub_status,
                    "error_count": len(sub_errs),
                    "errors": sub_errs,
                    "findings": sub_findings
                }]

                sub_obj = {
                    "section_id": sub_sec_id,
                    "section_name": sub_name,
                    "checks": [{"check_name": "Heading", "validation_results": sub_val_results}]
                }
                sec_obj["subsections"].append(sub_obj)

            structured_sections.append(sec_obj)
            continue  # skip the generic sec_obj append below

        else:
            # All other sections: subsections stay as checklist_name rows
            for sub_key, sub_info in h2_dict.items():
                sub_name = sub_info.get("heading") or sub_info.get("expected") or sub_key
                sub_errs = sub_errors.get(sub_key, [])
                sub_status = "FAIL" if sub_errs else "PASS"
                sub_findings = "Issues found." if sub_errs else "No findings."

                val_results.append({
                    "checklist_name": sub_name,
                    "status": sub_status,
                    "error_count": len(sub_errs),
                    "errors": sub_errs,
                    "findings": sub_findings
                })

                # Nested test scenarios
                for ts_name in (sub_info.get("test_scenarios") or []):
                    val_results.append({
                        "checklist_name": f"  \u2022 {ts_name}",
                        "status": sub_status,
                        "error_count": 0,
                        "errors": [],
                        "findings": "Verified." if sub_status == "PASS" else "Check parent section."
                    })

        sec_obj = {
            "section_id": f"section_{i+1}",
            "section_name": sec_name,
            "checks": [
                {
                    "check_name": "Heading",
                    "validation_results": val_results
                }
            ]
        }
        structured_sections.append(sec_obj)
    # Determine overall document validation status
    # Determine overall document validation status
    overall_val_status = validator_result.get("summary", {}).get("status", "PASS")

    structured_output = {
        "checks": [
            {
                "check_name": "Heading",
                "total_checklist_name": [f"Section Structure & Completeness - {overall_val_status}"]
            }
        ],
        "sections": structured_sections
    }

    output_json_path = output_dir / "output.json"
    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(structured_output, f, indent=2, ensure_ascii=False)
    print(f"  Saved: {output_json_path} (Structured Summary)")

    # ==================================================================
    # STEP 2b2 — Build Hierarchical Checklist (checklist_output.json)
    # ==================================================================
    checklist_output = build_checklist_output(validator_result, lossless_data)
    checklist_output_path = output_dir / "checklist_output.json"
    with open(checklist_output_path, "w", encoding="utf-8") as f:
        json.dump(checklist_output, f, indent=2, ensure_ascii=False)
    print(f"  Saved: {checklist_output_path} (Hierarchical Checklist)")

    # ==================================================================
    # STEP 2c — Build Pipeline Template Skeleton (pipeline_output.json)
    # ==================================================================
    skeleton_sections = []
    SEC8_SUB_IDS = {
        "sec8_1": "section_8_1",
        "sec8_2": "section_8_2",
        "sec8_3": "section_8_3",
        "sec8_4": "section_8_4",
    }

    # section_0 — Front Page skeleton (always first)
    skeleton_sections.append({
        "section_id":   "section_0",
        "section_name": "Front Page",
        "checks": [
            {
                "check_name": "",
                "validation_results": [
                    {
                        "checklist_name": "",
                        "status":         "",
                        "error_count":    0,
                        "errors":         [],
                        "findings":       ""
                    }
                ]
            }
        ]
    })

    for i, (sec_key, sec_name) in enumerate(_EXPECTED_SECTIONS):
        # Base skeleton for a section
        sec_skeleton = {
            "section_id": f"section_{i+1}",
            "section_name": sec_name,
            "checks": [
                {
                    "check_name": "",
                    "validation_results": [
                        {
                            "checklist_name": "",
                            "status": "",
                            "error_count": 0,
                            "errors": [],
                            "findings": ""
                        }
                    ]
                }
            ]
        }
        
        # Handle nesting for 8 and 11 to match output.json
        if sec_key in ("sec8", "sec11"):
            sec_skeleton["subsections"] = []
            p1_sec = validator_result.get("phase1", {}).get("sections", {}).get(sec_key, {})
            tree = p1_sec.get("heading_tree", {})
            h2_dict = tree.get("h2_children") or tree.get("h2_test_cases") or {}
            
            SEC8_EXPECTED = [
                ("sec8_1", "8.1. Number of Test Scenarios", "section_8_1"),
                ("sec8_2", "8.2. Test Bed Diagram", "section_8_2"),
                ("sec8_3", "8.3. Tools Required", "section_8_3"),
                ("sec8_4", "8.4. Test Execution Steps", "section_8_4"),
            ]

            if sec_key == "sec8":
                items = SEC8_EXPECTED
            else:
                items = []
                for sub_key, sub_info in h2_dict.items():
                    # h2_test_cases entries store the actual heading text under "found"
                    sub_name = sub_info.get("found") or sub_info.get("heading") or sub_info.get("expected") or sub_key
                    sub_id = f"section_11_{sub_key.split('_')[-1]}" if sub_key.startswith("tc_") else f"section_{sub_key}"
                    items.append((sub_key, sub_name, sub_id))

            for sub_key, sub_name, sub_sec_id in items:
                sub_skeleton = {
                    "section_id": sub_sec_id,
                    "section_name": sub_name,
                    "checks": [
                        {
                            "check_name": "",
                            "validation_results": [
                                {
                                    "checklist_name": "",
                                    "status": "",
                                    "error_count": 0,
                                    "errors": [],
                                    "findings": ""
                                }
                            ]
                        }
                    ]
                }
                sec_skeleton["subsections"].append(sub_skeleton)
        
        skeleton_sections.append(sec_skeleton)

    pipeline_output = {
        "skeleton": {
            "checks": [
                {
                    "check_name": "",
                    "total_checklist_name": [""]
                }
            ],
            "sections": skeleton_sections
        }
    }
    
    pipeline_output_path = output_dir / "pipeline_output.json"
    with open(pipeline_output_path, "w", encoding="utf-8") as f:
        json.dump(pipeline_output, f, indent=2, ensure_ascii=False)
    print(f"  Saved: {pipeline_output_path} (Skeleton Template)")

    # Summarise per-section validation results
    sections_result   = validator_result.get("sections") or {}
    split_map_detail  = validator_result.get("split_map_detail") or {}
    validator_summary = validator_result.get("summary") or {}

    warned_keys:  List[str] = []
    passing_keys: List[str] = []

    for sec_key, sec_info in sections_result.items():
        validation = (sec_info.get("validation") or "").upper()
        if sec_key not in failed_keys:
            if validation == "WARN":
                warned_keys.append(sec_key)
            else:
                passing_keys.append(sec_key)

    print(f"\n  Validator summary:")
    print(f"    Overall status : {validator_summary.get('status', 'N/A')}")
    print(f"    Sections pass  : {len(passing_keys)}")
    print(f"    Sections warn  : {len(warned_keys)}")
    print(f"    Sections fail  : {len(failed_keys)}")

    if failed_keys:
        print(f"\n  [!] FAILED sections (will be SKIPPED from structured output):")
        for sec_key in sorted(failed_keys):
            sec_num   = _SEC_KEY_TO_NUM.get(sec_key, sec_key)
            sec_info  = sections_result.get(sec_key, {})
            sec_title = sec_info.get("section", sec_key)
            issues    = sec_info.get("issues") or []
            print(f"      Section {sec_num} — {sec_title}  ({len(issues)} issue(s))")
            for iss in issues[:5]:
                what = (iss.get("what") or "")[:120]
                print(f"        • {what}")

    if warned_keys:
        print(f"\n  [~] WARN sections (included in output with warnings):")
        for sec_key in sorted(warned_keys):
            sec_num   = _SEC_KEY_TO_NUM.get(sec_key, sec_key)
            sec_title = sections_result[sec_key].get("section", sec_key)
            print(f"      Section {sec_num} — {sec_title}")

    # ==================================================================
    # STEP 3 — Remove failed sections from the mapped_sections map
    # ==================================================================
    # Failed section keys are simply deleted from mapped_sections.
    # build_from_map() will then skip them entirely — no block filtering needed.

    # ==================================================================
    # STEP 4 — Build structured JSON using mapped_sections (no bleed)
    # ==================================================================
    print(f"\n{'='*60}")
    print(f"STEP 3 — Building structured JSON (map-driven)")
    print(f"{'='*60}")

    # Build a clean mapped_sections with failed keys removed
    mapped_sections: Dict[str, Any] = dict(validator_result.get("mapped_sections", {}))
    for fk in failed_keys:
        mapped_sections.pop(fk, None)
    # If all sec11 test cases failed, drop tc_* keys too
    if "sec11" not in mapped_sections:
        for k in list(mapped_sections.keys()):
            if k.startswith("tc_"):
                mapped_sections.pop(k, None)

    # Build a heading-text lookup for every failed tc_* key so the FAIL stubs
    # in the structured / ai_structured JSONs can show the real document title
    # (e.g. "11.1.7 Test Case Number:") instead of a generic placeholder.
    _split_map_detail = validator_result.get("split_map_detail") or {}
    failed_tc_headings: Dict[str, str] = {
        fk: (_split_map_detail.get(fk) or {}).get("heading") or fk
        for fk in failed_keys
        if fk.startswith("tc_")
    }

    try:
        builder         = StructuredSectionBuilder(lossless_data["document"])
        structured_data = builder.build_from_map(
            mapped_sections,
            failed_keys=set(failed_keys),
            failed_tc_headings=failed_tc_headings,
        ).to_dict()
    except Exception as exc:
        print(f"[ERROR] structured_extract (build_from_map) failed: {exc}")
        structured_data = {"document": lossless_data.get("document"), "sections": [], "error": str(exc)}

    # If sec1 heading is missing, the front page end-boundary is unknown;
    # its content may be contaminated — mark FAIL and strip the content entirely.
    if "sec1" in failed_keys and structured_data.get("frontpage_data"):
        structured_data["frontpage_data"] = {
            "section_id": "FP-01",
            "status": "FAIL",
            "error": "Front Page end-boundary unknown — Section 1 heading is missing.",
        }

    structured_path = output_dir / f"{base_name}_structured.json"
    with open(structured_path, "w", encoding="utf-8") as f:
        json.dump(structured_data, f, indent=2, ensure_ascii=False)
    print(f"  Saved: {structured_path}")
    print(f"  Sections in output: {len(structured_data.get('sections', []))}")

    # ==================================================================
    # STEP 5 — Build AI structured JSON using mapped_sections (no bleed)
    # ==================================================================
    print(f"\n{'='*60}")
    print(f"STEP 4 — Building AI structured JSON (map-driven)")
    print(f"{'='*60}")
    try:
        ai_builder         = AIStructuredSectionBuilder(lossless_data["document"])
        ai_structured_data = ai_builder.build_from_map(
            mapped_sections,
            failed_keys=set(failed_keys),
            failed_tc_headings=failed_tc_headings,
        ).to_dict()
    except Exception as exc:
        print(f"[ERROR] AI_structured_extract (build_from_map) failed: {exc}")
        ai_structured_data = {"document": lossless_data.get("document"), "sections": [], "error": str(exc)}

    # Same frontpage boundary check for AI-structured output
    if "sec1" in failed_keys and ai_structured_data.get("frontpage_data"):
        ai_structured_data["frontpage_data"] = {
            "section_id": "FP-01",
            "status": "FAIL",
            "error": "Front Page end-boundary unknown — Section 1 heading is missing.",
        }

    ai_structured_path = output_dir / f"{base_name}_ai_structured.json"
    with open(ai_structured_path, "w", encoding="utf-8") as f:
        json.dump(ai_structured_data, f, indent=2, ensure_ascii=False)
    print(f"  Saved: {ai_structured_path}")
    print(f"  Sections in output: {len(ai_structured_data.get('sections', []))}")
    # ==================================================================
    # FINAL SUMMARY
    # ==================================================================
    print(f"\n{'='*60}")
    print(f"PIPELINE COMPLETE")
    print(f"{'='*60}")
    print(f"  Input file      : {file_path}")
    print(f"  Lossless JSON   : {lossless_path}")
    print(f"  Validator output: {output_json_path}")
    print(f"  Structured JSON : {structured_path}")
    print(f"  AI Structured   : {ai_structured_path}")

    if failed_keys:
        skipped_nums = sorted(
            str(_SEC_KEY_TO_NUM.get(k, k)) for k in failed_keys
        )
        print(f"\n  SKIPPED sections (validation FAIL): {skipped_nums}")
        print(f"  These sections had errors and were excluded from")
        print(f"  structured.json and ai_structured.json.")
        included_nums = sorted(
            str(_SEC_KEY_TO_NUM.get(k, k))
            for k in list(passing_keys) + list(warned_keys)
        )
        print(f"  Included sections: {included_nums}")
    else:
        print(f"\n  All sections passed — nothing skipped.")

    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()