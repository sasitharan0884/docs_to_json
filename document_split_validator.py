"""
document_split_validator.py
============================
Multi-Phase DOCX Validator — Split & Analyze

How it works
------------
  Step 0  Parse lossless JSON → build block array
  Step 0b Build Split Map  — map every section/subsection to its block_id range
            {"front_page": "0-43", "sec1": "44-46", ...}

  Phase 1  Heading Structure Validation (per section, non-fatal)
            - Checks all 12 H1s exist and have correct number prefixes
            - Checks Section 8 has all 4 H2 children (8.1 / 8.2 / 8.3 / 8.4)
            - Checks Section 11 every H2 test-case has all 5 H3 subsections (a-e)
            - Checks Section 9, 11, 12 scenario count consistency
            Each section gets its own status: PASS | WARN | FAIL
            Validation NEVER stops — all sections are always checked.

  Phase 2  Content Extraction (independent per section, continues on errors)
            Uses the structured extractors from AI_structured_extract.py.
            Even if a section has Phase-1 warnings/failures, extraction is still
            attempted — the section is marked EXTRACTED or EXTRACTION_FAILED.

  Output   output.json with:
            {
              "split_map":   { section_key: {start_bid, end_bid, range, heading} },
              "phase1":      { summary, sections: { ... } },
              "phase2":      { summary, sections: { ... } },
            }

Usage
-----
  # From a DOCX (auto-runs lossless extraction):
  python document_split_validator.py input.docx

  # From an existing lossless JSON:
  python document_split_validator.py lossless.json

  # Explicit output path:
  python document_split_validator.py input.docx --out my_output.json
"""

from __future__ import annotations
import json
import re
import sys
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# CONSTANTS — expected document schema
# ---------------------------------------------------------------------------

EXPECTED_H1 = [
    {"key": "sec1",  "num": "1",  "name": "ITSAR Section No & Name"},
    {"key": "sec2",  "num": "2",  "name": "Security Requirement No & Name"},
    {"key": "sec3",  "num": "3",  "name": "Requirement Description"},
    {"key": "sec4",  "num": "4",  "name": "DUT Confirmation Details"},
    {"key": "sec5",  "num": "5",  "name": "DUT Configuration"},
    {"key": "sec6",  "num": "6",  "name": "Preconditions"},
    {"key": "sec7",  "num": "7",  "name": "Test Objective"},
    {"key": "sec8",  "num": "8",  "name": "Test Plan"},
    {"key": "sec9",  "num": "9",  "name": "Expected Results for Pass"},
    {"key": "sec10", "num": "10", "name": "Expected Format of Evidence"},
    {"key": "sec11", "num": "11", "name": "Test Execution"},
    {"key": "sec12", "num": "12", "name": "Test Case Result"},
]

EXPECTED_SEC8_H2 = [
    {"key": "sec8_1", "num": "8.1", "name": "Number of Test Scenarios"},
    {"key": "sec8_2", "num": "8.2", "name": "Test Bed Diagram"},
    {"key": "sec8_3", "num": "8.3", "name": "Tools Required"},
    {"key": "sec8_4", "num": "8.4", "name": "Test Execution Steps"},
]

EXPECTED_SEC11_H3 = [
    {"key": "a", "name": "Test Case Name"},
    {"key": "b", "name": "Test Case Description"},
    {"key": "c", "name": "Execution Steps"},
    {"key": "d", "name": "Test Observations"},
    {"key": "e", "name": "Evidence Provided"},
]

SCENARIO_HEADER_RE = re.compile(
    r"^\s*test\s*(scenario|case)s?\s+\d+(?:\.\d+){2,}\b",
    re.IGNORECASE,
)

# Matches bare "Test Scenario: ..." or "Test Case: ..." with no dotted number.
_SCENARIO_BARE_RE = re.compile(
    r"^\s*test\s*(?:scenario|case)s?\s*:\s*",
    re.IGNORECASE,
)


def _is_bold_scenario_header(text: str, bold: str) -> bool:
    """Return True if *text* opens a new scenario block and is formatted bold."""
    if not bold:
        return False
    if SCENARIO_HEADER_RE.match(text):
        return True
    if _SCENARIO_BARE_RE.match(text):
        return True
    return False

# ---------------------------------------------------------------------------
# NORMALISATION HELPERS
# ---------------------------------------------------------------------------

def _norm_alpha(text: str) -> str:
    """Keep only letters, lowercase — for strict name comparison."""
    return re.sub(r"[^a-zA-Z]+", "", text or "").lower()


def _norm_name(text: str) -> str:
    """Strip leading numbering, lower, alphanum-only — for fuzzy name match."""
    t = re.sub(r"^\s*\d+(?:\.\d+)*\s*[.):]?\s*", "", (text or "").strip())
    t = t.strip().lower()
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _strip_alpha_prefix(text: str) -> str:
    """Strip a single leading letter prefix (a., b., c.) from text.
    
    e.g. 'a. Test Case Name: ' -> 'Test Case Name: '
         'b. Test Case Description:' -> 'Test Case Description:'
    """
    return re.sub(r"^\s*[a-zA-Z]\s*[.):]\s*", "", (text or "").strip()).strip()


def _match_h3_name(text: str, expected_name: str) -> bool:
    """Match an H3 heading text against an expected subsection name.
    
    Handles headings prefixed with 'a.', 'b.', 'c.' etc.
    e.g. 'a. Test Case Name: ' matches 'Test Case Name'
    """
    # Strip letter prefix from the actual heading text, then compare
    stripped = _strip_alpha_prefix(text)
    return (
        _norm_alpha(stripped) == _norm_alpha(expected_name)
        or _norm_name(stripped) == _norm_name(expected_name)
        or _match_name(text, expected_name)  # fallback to original logic
    )



def _match_name(text: str, expected_name: str) -> bool:
    """True if *text* refers to the same section as *expected_name*."""
    return (
        _norm_alpha(text) == _norm_alpha(expected_name)
        or _norm_name(text) == _norm_name(expected_name)
    )


def _has_number_prefix(text: str, num: str) -> bool:
    """True if *text* starts with the given numeric prefix (e.g. '8.')."""
    return bool(re.match(rf"^\s*{re.escape(num)}[.\s:]", text or ""))


# ---------------------------------------------------------------------------
# STEP 0b — SPLIT MAP BUILDER
# ---------------------------------------------------------------------------

class SplitMap:
    """
    Implements the 5-step Master Design for document splitting:
    PHASE A: Lossless Block Input (via __init__)
    PHASE B: H1 Splitter Engine (Top-level sections 1-12)
    PHASE C: Nested Section Splitter (8.1-8.4 and Section 11 Test Cases)
    """

    def __init__(self, blocks: List[Dict[str, Any]]) -> None:
        self.blocks = blocks
        # Step 2: Store all splits in dictionary
        self.section_dict: Dict[str, Any] = {}
        # Backwards compatibility: keep .map for existing calls
        self.map: Dict[str, Any] = {} 
        
        self._phase_b_h1_splitter()
        self._phase_c_nested_splitter()
        
        # Sync .map for legacy code (Phase 2 extractors etc)
        self._sync_legacy_map()

    def _phase_b_h1_splitter(self) -> None:
        """Step 1 & 2: Split document into top-level section records."""
        
        # Initialize dictionary with all expected sections as MISSING
        self.section_dict = {
            "front_page": {
                "expected_name": "Front Matter",
                "found_heading": "[Front Page]",
                "status": "PASS",
                "start_bid": self.blocks[0]["block_id"] if self.blocks else 0,
                "end_bid": 0,
                "idx_start": 0,
                "idx_end": -1,
                "blocks": []
            }
        }
        for es in EXPECTED_H1:
            self.section_dict[es["key"]] = {
                "expected_name": es["name"],
                "expected_num": es["num"],
                "found_heading": None,
                "status": "MISSING",
                "start_bid": None,
                "end_bid": None,
                "idx_start": -1,
                "idx_end": -1,
                "blocks": []
            }

        # Step 1: Find all H1 headings
        matches: Dict[str, List[Dict]] = {es["key"]: [] for es in EXPECTED_H1}
        for idx, b in enumerate(self.blocks):
            if b.get("style") == "Heading 1":
                text = (b.get("text") or "").strip()
                key = self._match_h1_key(text)
                if key in matches:
                    matches[key].append({"idx": idx, "bid": b["block_id"], "text": text})

        # Step 2: Populate section_dict and collect for boundary calculation
        h1_entries = []
        for key, found in matches.items():
            if not found:
                continue
            
            # AMBIGUOUS if multiple H1s match the same section
            status = "PASS" if len(found) == 1 else "AMBIGUOUS"
            
            self.section_dict[key].update({
                "found_heading": found[0]["text"],
                "status": status,
                "start_bid": found[0]["bid"],
                "idx_start": found[0]["idx"],
                "occurrences": len(found) if status == "AMBIGUOUS" else 1
            })
            h1_entries.append(found[0])

        # Step 3: Determine boundaries & Apply "BOUNDARY_BROKEN" rule
        h1_entries.sort(key=lambda x: x["idx"])
        
        # Front page boundary
        if h1_entries:
            fp_end_idx = h1_entries[0]["idx"] - 1
            self.section_dict["front_page"]["idx_end"] = max(fp_end_idx, 0)
            self.section_dict["front_page"]["end_bid"] = self.blocks[self.section_dict["front_page"]["idx_end"]]["block_id"] if self.blocks else 0
        else:
            self.section_dict["front_page"]["idx_end"] = len(self.blocks) - 1
            self.section_dict["front_page"]["end_bid"] = self.blocks[-1]["block_id"] if self.blocks else 0

        # Section boundaries
        for i, entry in enumerate(h1_entries):
            key = self._match_h1_key(entry["text"])
            if key not in self.section_dict: continue
            
            next_idx = h1_entries[i+1]["idx"] if i+1 < len(h1_entries) else len(self.blocks)
            self.section_dict[key]["idx_end"] = next_idx - 1
            self.section_dict[key]["end_bid"] = self.blocks[next_idx - 1]["block_id"]

        # Step 3: Validate dictionary completeness (Boundary Rule)
        for i, es in enumerate(EXPECTED_H1):
            key = es["key"]
            if i + 1 < len(EXPECTED_H1):
                next_key = EXPECTED_H1[i+1]["key"]
                if self.section_dict[next_key]["status"] == "MISSING" and self.section_dict[key]["status"] in ("PASS", "AMBIGUOUS"):
                    self.section_dict[key]["status"] = "BOUNDARY_BROKEN"

        # Finalize Phase B blocks
        for key, info in self.section_dict.items():
            if info["idx_start"] != -1 and info["idx_end"] != -1:
                info["blocks"] = self.blocks[info["idx_start"] : info["idx_end"] + 1]

    def _phase_c_nested_splitter(self) -> None:
        """Step 4: Process nested sections (8 / 11)."""
        
        # Section 8 Sub-splitter
        sec8 = self.section_dict.get("sec8")
        if sec8 and sec8["status"] != "MISSING":
            self._split_sec8(sec8)
            
        # Section 11 Sub-splitter (Test Cases)
        sec11 = self.section_dict.get("sec11")
        if sec11 and sec11["status"] != "MISSING":
            self._split_sec11(sec11)

    def _split_sec8(self, parent: Dict) -> None:
        """PHASE C — SECTION 8 SUB-SPLITTER"""
        blocks = parent["blocks"]
        idx_offset = parent["idx_start"]
        
        # Find H2s within Section 8 (Phase 10 splitting only)
        h2_found = []
        for i, b in enumerate(blocks):
            style = b.get("style", "")
            text = (b.get("text") or "").strip()
            norm = _norm_alpha(text)
            key = self._match_sec8_h2_key(text)
            
            if key or style == "Heading 2":
                is_styled = (style == "Heading 2")
                h2_found.append({
                    "idx_in_parent": i, "bid": b["block_id"], 
                    "key": key, "text": text, "is_styled": is_styled
                })

        # Correct Boundary Update: sec8 parent content ends before 8.1 starts
        if h2_found:
            first_h2_idx = h2_found[0]["idx_in_parent"]
            if first_h2_idx > 0:
                parent["idx_end"] = idx_offset + first_h2_idx - 1
                parent["end_bid"] = self.blocks[parent["idx_end"]]["block_id"]
                parent["blocks"] = self.blocks[parent["idx_start"] : parent["idx_end"] + 1]

        # Process each expected H2 (8.1 - 8.4)
        for i, es8 in enumerate(EXPECTED_SEC8_H2):
            key = es8["key"]
            match = next((h for h in h2_found if h["key"] == key), None)
            
            info = {
                "expected_name": es8["name"],
                "expected_num": es8["num"],
                "found_heading": match["text"] if match else None,
                "status": "PASS" if (match and match["is_styled"]) else "MISSING",
                "is_styled": match["is_styled"] if match else False,
                "start_bid": match["bid"] if match else None,
                "end_bid": None,
                "idx_start": idx_offset + match["idx_in_parent"] if match else -1,
                "idx_end": -1,
                "blocks": [],
                "parent": "sec8"
            }
            
            if match:
                m_idx = next(j for j, h in enumerate(h2_found) if h["key"] == key)
                next_h2_idx_in_parent = h2_found[m_idx + 1]["idx_in_parent"] if m_idx + 1 < len(h2_found) else len(blocks)
                
                info["idx_end"] = idx_offset + next_h2_idx_in_parent - 1
                info["end_bid"] = self.blocks[info["idx_end"]]["block_id"]
                info["blocks"] = self.blocks[info["idx_start"] : info["idx_end"] + 1]
                
                # Boundary uncertainty within Section 8
                if i + 1 < len(EXPECTED_SEC8_H2):
                    next_expected_key = EXPECTED_SEC8_H2[i+1]["key"]
                    if not any(h["key"] == next_expected_key for h in h2_found):
                        info["status"] = "BOUNDARY_BROKEN"
            
            self.section_dict[key] = info

        # If any H2 missing, parent is INCOMPLETE
        if any(self.section_dict[es8["key"]]["status"] == "MISSING" for es8 in EXPECTED_SEC8_H2):
            if parent["status"] == "PASS": parent["status"] = "INCOMPLETE"

    def _split_sec11(self, parent: Dict) -> None:
        """
        Section 11 Production Engine — Phase 1 & 2 & 10.
        Splits Section 11 blocks into test-case groups (tc_1, tc_2, etc).
        Includes recovery for styled-missing headings.
        """
        blocks = parent["blocks"]
        idx_offset = parent["idx_start"]
        
        # Phase 1 & 10: Normalize and Detect H2/H3 markers
        tokens = []
        for i, b in enumerate(blocks):
            style = b.get("style", "")
            text = (b.get("text") or "").strip()
            norm = _norm_alpha(text)
            
            token_type = "NORMAL"
            h3_key = None
            
            # H2 Detection (Production Rules)
            is_styled = False
            if style == "Heading 2" and norm.startswith("testcasenumber"):
                token_type = "H2_TC"
                is_styled = True
            elif norm.startswith("testcasenumber"):
                token_type = "H2_TC" # Phase 10 Recovery
                is_styled = False
                
            # H3 Detection (Production Rules)
            elif style == "Heading 3" or True: # Check all for Phase 10 recovery
                for pattern, key in {
                    "atestcasename": "a",
                    "testcasename": "a",
                    "btestcasedescription": "b",
                    "testcasedescription": "b",
                    "cexecutionsteps": "c",
                    "executionsteps": "c",
                    "dtestobservations": "d",
                    "testobservations": "d",
                    "dtestobservation": "d",
                    "testobservation": "d",
                    "eevidenceprovided": "e",
                    "evidenceprovided": "e"
                }.items():
                    if norm.startswith(pattern):
                        token_type = f"H3_{key.upper()}"
                        h3_key = key
                        is_styled = (style == "Heading 3")
                        break
            
            if b.get("type") == "table": token_type = "TABLE"
            if b.get("type") == "image": token_type = "IMAGE"

            tokens.append({
                "idx_in_parent": i,
                "token_type": token_type,
                "h3_key": h3_key,
                "text": text,
                "bid": b["block_id"],
                "is_styled": is_styled
            })

        # Phase 2: Split by H2
        tc_groups = []
        current_tc = None
        
        for t in tokens:
            if t["token_type"] == "H2_TC":
                if current_tc: tc_groups.append(current_tc)
                current_tc = {"heading_token": t, "tokens": [t]}
            elif current_tc:
                current_tc["tokens"].append(t)
            elif t["token_type"].startswith("H3_"):
                # Phase 3/Edge Case A: No H2 but H3 found -> Orphan
                current_tc = {"heading_token": None, "tokens": [t], "orphan": True}

        if current_tc: tc_groups.append(current_tc)

        if tc_groups:
            # Truncate parent at first TC
            first_tc_idx = tc_groups[0]["tokens"][0]["idx_in_parent"]
            if first_tc_idx > 0:
                parent["idx_end"] = idx_offset + first_tc_idx - 1
                parent["end_bid"] = self.blocks[parent["idx_end"]]["block_id"]
                parent["blocks"] = self.blocks[parent["idx_start"] : parent["idx_end"] + 1]

        # Phase 6 & 7: Hidden Testcase Recovery (Recursive)
        # If inside a TC tokens, after 'e', another 'a' starts, split it.
        final_tc_groups = []
        for group in tc_groups:
            sub_groups = self._recursive_split_hidden_tc(group)
            final_tc_groups.extend(sub_groups)

        for j, group in enumerate(final_tc_groups):
            tc_key = f"tc_{j + 1}"
            tokens = group["tokens"]
            idx_start = idx_offset + tokens[0]["idx_in_parent"]
            idx_end = idx_offset + tokens[-1]["idx_in_parent"]
            
            is_styled = group["heading_token"]["is_styled"] if group["heading_token"] else False
            status = "PASS" if is_styled else "FAIL"
            if group.get("orphan"): status = "FAIL"
            if len(tokens) == 1 and tokens[0]["token_type"] == "H2_TC": status = "FAIL" # Empty TC

            self.section_dict[tc_key] = {
                "expected_name": f"Test Case {j+1}",
                "found_heading": group["heading_token"]["text"] if group["heading_token"] else "A properly formatted 'Heading 2' test case number is missing before the test case content.",
                "status": status,
                "start_bid": tokens[0]["bid"],
                "end_bid": self.blocks[idx_end]["block_id"],
                "idx_start": idx_start,
                "idx_end": idx_end,
                "blocks": self.blocks[idx_start : idx_end + 1],
                "parent": "sec11",
                "orphan": group.get("orphan", False),
                "is_styled": is_styled,
                "production_tokens": tokens # Store for Phase 3 validation
            }

    def _recursive_split_hidden_tc(self, group: Dict) -> List[Dict]:
        """Phase 6 & 7: Recovery logic for hidden testcases after 'e'."""
        tokens = group["tokens"]
        if not tokens: return []
        
        # Find 'e' and check if 'a' restarts
        e_pos = -1
        for i, t in enumerate(tokens):
            if t["token_type"] == "H3_E":
                e_pos = i
                # Note: don't break, find last 'e' if duplicates exist? 
                # Actually, according to Phase 5, after 'e', if 'a' starts...
        
        if e_pos == -1 or e_pos == len(tokens) - 1:
            return [group]

        # Check for restart of 'a' after 'e'
        restart_pos = -1
        for i in range(e_pos + 1, len(tokens)):
            if tokens[i]["token_type"] == "H3_A":
                restart_pos = i
                break
        
        if restart_pos == -1:
            return [group]

        # Split!
        head = {"heading_token": group["heading_token"], "tokens": tokens[:restart_pos], "orphan": group.get("orphan", False)}
        tail_tokens = tokens[restart_pos:]
        
        # Check if tail has a hidden H2 (Phase 10)
        tail_heading = None
        for t in tail_tokens:
            if t["token_type"] == "H2_TC":
                tail_heading = t
                break
        
        tail = {"heading_token": tail_heading, "tokens": tail_tokens, "orphan": tail_heading is None}
        
        # Recurse on tail
        return [head] + self._recursive_split_hidden_tc(tail)

    def _match_h1_key(self, text: str) -> Optional[str]:
        for es in EXPECTED_H1:
            if _match_name(text, es["name"]):
                return es["key"]
        return None

    def _match_sec8_h2_key(self, text: str) -> Optional[str]:
        for es8 in EXPECTED_SEC8_H2:
            if _match_name(text, es8["name"]):
                return es8["key"]
        return None

    def _sync_legacy_map(self) -> None:
        """Populate the old .map for compatibility."""
        for key, info in self.section_dict.items():
            self.map[key] = {
                "start_bid": info["start_bid"],
                "end_bid": info["end_bid"],
                "range": f"{info['start_bid']}-{info['end_bid']}" if info["start_bid"] is not None else "N/A",
                "heading": info["found_heading"] or "[Missing]",
                "level": 1 if key.startswith("sec") and "_" not in key else 2 if "_" in key or key.startswith("tc_") else 0,
                "idx_start": info["idx_start"],
                "idx_end": info["idx_end"],
            }

    def slice(self, key: str) -> List[Dict[str, Any]]:
        """Return the block slice for a given section key."""
        info = self.section_dict.get(key)
        if not info:
            return []
        return info["blocks"]

    def to_display(self) -> Dict[str, str]:
        """Return { key: range_string } for JSON output."""
        return {k: v["range"] for k, v in self.map.items()}


# ---------------------------------------------------------------------------
# PHASE 1 — HEADING STRUCTURE VALIDATION
# ---------------------------------------------------------------------------

def _make_issue(itype: str, severity: str, message: str, suggestion: str = "", where: str = "", what: str = "", redirect_text: str = "") -> Dict:
    res = {
        "type": itype,
        "severity": severity,
        "message": message,
        "suggestion": suggestion,
    }
    if where: res["where"] = where
    if what: res["what"] = what
    if redirect_text: res["redirect_text"] = redirect_text
    return res


def validate_phase1(blocks: List[Dict], split_map: SplitMap) -> Dict[str, Any]:
    """
    Non-fatal per-section heading validation.
    Returns { summary, sections }.
    """
    sections: Dict[str, Any] = {}

    # ---- Validate each expected H1 section ----
    for i, es in enumerate(EXPECTED_H1):
        key = es["key"]
        num = es["num"]
        name = es["name"]
        issues: List[Dict] = []
        heading_tree: Dict[str, Any] = {}

        # Master Design: Get section from splitter's section_dict
        sec_info = split_map.section_dict.get(key)
        if not sec_info:
            continue

        found_heading = sec_info["found_heading"]
        status_from_splitter = sec_info["status"]

        # 1. H1 presence
        if status_from_splitter == "MISSING":
            full_name = f"{num}. {name}"
            issues.append(_make_issue(
                "MISSING_H1", "HIGH",
                f"Heading 1 for section '{full_name}' was not found.",
                f"Add the section using the exact title '{full_name}' and apply Heading 1 formatting.",
                where=full_name,
                what=f"The required section '{full_name}' is missing or not formatted using Heading 1 style.",
                redirect_text=full_name
            ))
            sections[key] = {
                "status": "FAIL",
                "expected": f"{num}. {name}",
                "found": None,
                "block_range": "N/A",
                "issues": issues,
                "heading_tree": heading_tree,
            }
            continue

        heading_tree["h1"] = {"text": found_heading, "status": "PASS"}

        # 2. Number prefix check
        if not _has_number_prefix(found_heading, num):
            full_name = f"{num}. {name}"
            issues.append(_make_issue(
                "MISSING_NUMBER_PREFIX", "WARN",
                f"Expected numbering '{num}.' before '{name}', found: '{found_heading}'.",
                f"Update the heading to start with '{num}. {name}'.",
                where=full_name,
                what=f"Section heading '{found_heading}' is missing the required '{num}.' number prefix.",
                redirect_text=full_name
            ))

        # 2b. Master Design: Boundary Uncertainty check
        if status_from_splitter == "BOUNDARY_BROKEN":
            full_name = f"{num}. {name}"
            next_num = EXPECTED_H1[i + 1]["num"] if i + 1 < len(EXPECTED_H1) else "?"
            next_name = EXPECTED_H1[i + 1]["name"] if i + 1 < len(EXPECTED_H1) else "Next Section"
            issues.append(_make_issue(
                "INVALID_BOUNDARY", "HIGH",
                f"The next expected section '{next_num}. {next_name}' is missing, so the end of this section cannot be reliably determined.",
                f"Ensure '{next_num}. {next_name}' is present so this section can be correctly bounded.",
                where=full_name,
                what=f"Section boundary for '{full_name}' is broken due to the absence of the subsequent section heading.",
                redirect_text=full_name
            ))

        # 2c. Ambiguity check
        if status_from_splitter == "AMBIGUOUS":
            full_name = f"{num}. {name}"
            issues.append(_make_issue(
                "AMBIGUOUS_HEADING", "WARN",
                f"Multiple Heading 1s matching '{name}' were found. Using the first occurrence at block {sec_info['start_bid']}.",
                "Remove duplicate headings so the structure is unique.",
                where=full_name,
                what=f"Found multiple Heading 1 paragraphs matching the '{full_name}' section criteria.",
                redirect_text=full_name
            ))

        # 3. Section-8 H2 children
        if key == "sec8":
            h2_children: Dict[str, Any] = {}
            for es8 in EXPECTED_SEC8_H2:
                sub_key = es8["key"]
                sub_info = split_map.section_dict.get(sub_key)
                if sub_info and sub_info["status"] != "MISSING":
                    h2_children[sub_key] = {
                        "status": sub_info["status"],
                        "expected": f"{es8['num']}. {es8['name']}",
                        "found": sub_info["found_heading"],
                        "block_range": f"{sub_info['start_bid']}-{sub_info['end_bid']}",
                        "num_prefix": _has_number_prefix(sub_info["found_heading"], es8["num"]),
                    }
                    if sub_info["status"] == "BOUNDARY_BROKEN":
                        full_sub_name = f"{es8['num']}. {es8['name']}"
                        issues.append(_make_issue(
                            "INVALID_SUB_BOUNDARY", "MEDIUM",
                            f"Subsection '{full_sub_name}' has an uncertain end boundary.",
                            "Ensure the following subsection exists to correctly bound this content.",
                            where=full_sub_name,
                            what=f"The end of subsection '{full_sub_name}' cannot be determined because the subsequent heading is missing.",
                            redirect_text=full_sub_name
                        ))
                else:
                    h2_children[sub_key] = {
                        "status": "FAIL",
                        "expected": f"{es8['num']}. {es8['name']}",
                        "found": None,
                    }
                    full_sub_name = f"{es8['num']}. {es8['name']}"
                    issues.append(_make_issue(
                        "MISSING_H2", "HIGH",
                        f"Heading 2 for subsection '{full_sub_name}' was not found.",
                        f"Add the subsection using the exact title '{full_sub_name}' and apply Heading 2 formatting.",
                        where=full_sub_name,
                        what=f"The required subsection '{full_sub_name}' is missing or not formatted using Heading 2 style.",
                        redirect_text=full_sub_name
                    ))
            heading_tree["h2_children"] = h2_children

        # 4. Section-11 H2 test-case groups + H3 a-e
        if key == "sec11":
            tc_keys = sorted(
                [k for k in split_map.map if k.startswith("tc_")],
                key=lambda k: int(k.split("_")[1]),
            )
            if not tc_keys:
                issues.append(_make_issue(
                    "MISSING_TEST_CASES", "HIGH",
                    "No Heading 2 test-case groups were found inside Section 11.",
                    "Add test cases as Heading 2 entries under '11. Test Execution'.",
                ))
            else:
                tc_tree: Dict[str, Any] = {}
                for tc_key in tc_keys:
                    tc_info = split_map.section_dict[tc_key]
                    tokens = tc_info.get("production_tokens", [])
                    
                    # Master Design: Recovered headings are FAIL (Missing style)
                    is_h2_ok = tc_info.get("is_styled", False)
                    tc_status = "PASS" if is_h2_ok else "FAIL"
                    tc_issues = []
                    
                    if not is_h2_ok:
                        issues.append(_make_issue(
                            "INVALID_HEADING_STYLE", "MEDIUM",
                            f"Test case heading '{tc_info['found_heading']}' is not styled as 'Heading 2'.",
                            "Apply the 'Heading 2' style to this test case number.",
                            where=tc_info['found_heading'],
                            what=f"Test case identifier '{tc_info['found_heading']}' is present but lacks 'Heading 2' styling.",
                            redirect_text=tc_info['found_heading']
                        ))

                    if tc_info.get("orphan"):
                        tc_status = "FAIL"
                        tc_issues.append("A properly formatted 'Heading 2' test case number is missing before the test case content.")

                    # Phase 3 & 4: Validate H3 Presence, Order, Duplicates
                    h3_sequence = [] # List of (key, token)
                    for t in tokens:
                        if t["token_type"].startswith("H3_"):
                            h3_sequence.append((t["h3_key"], t))

                    found_keys = [s[0] for s in h3_sequence]
                    expected_keys = ["a", "b", "c", "d", "e"]
                    
                    h3_tree: Dict[str, Any] = {}
                    for ek in expected_keys:
                        matches = [s[1] for s in h3_sequence if s[0] == ek]
                        # PASS only if found AND correctly styled as Heading 3
                        is_ok = matches and matches[0].get("is_styled", False)
                        
                        h3_tree[ek] = {
                            "status": "PASS" if is_ok else "FAIL",
                            "expected": next(es["name"] for es in EXPECTED_SEC11_H3 if es["key"] == ek),
                            "found": matches[0]["text"] if matches else None,
                        }
                        
                        if not is_ok:
                            tc_status = "FAIL"
                            if not matches:
                                issues.append(_make_issue(
                                    "MISSING_H3", "MEDIUM",
                                    f"Test case '{tc_info['found_heading']}' is missing H3 subsection '{h3_tree[ek]['expected']}'.",
                                    f"Add Heading 3 '{ek}. {h3_tree[ek]['expected']}' inside this test case."
                                ))
                            else:
                                issues.append(_make_issue(
                                    "INVALID_H3_STYLE", "MEDIUM",
                                    f"Test case '{tc_info['found_heading']}' subsection '{h3_tree[ek]['expected']}' is not styled as 'Heading 3'.",
                                    f"Apply 'Heading 3' style to '{matches[0]['text']}'."
                                ))
                        elif len(matches) > 1:
                            issues.append(_make_issue(
                                "DUPLICATE_H3", "WARN",
                                f"Test case '{tc_info['found_heading']}' has duplicate H3 for '{h3_tree[ek]['expected']}'.",
                                "Remove redundant Heading 3 entries."
                            ))

                    # Phase 4 Case 4: Wrong Order
                    # Get first appearance index of each required key that was found
                    first_indices = []
                    for ek in expected_keys:
                        try:
                            idx = found_keys.index(ek)
                            first_indices.append(idx)
                        except ValueError:
                            pass
                    
                    if first_indices != sorted(first_indices):
                        tc_status = "FAIL"
                        issues.append(_make_issue(
                            "H3_OUT_OF_ORDER", "HIGH",
                            f"Test case '{tc_info['found_heading']}' has H3 subsections in the wrong order. Expected: a→b→c→d→e.",
                            "Reorder Heading 3 entries to follow the standard template."
                        ))

                    # Phase 8: Empty Testcase
                    if len(tokens) == 1 and tokens[0]["token_type"] == "H2_TC":
                        tc_status = "FAIL"
                        issues.append(_make_issue(
                            "EMPTY_TEST_CASE", "HIGH",
                            f"Test case '{tc_info['found_heading']}' has no content (H2 only).",
                            "Add test case details (Name, Description, Steps, etc.)."
                        ))

                    tc_tree[tc_key] = {
                        "found": tc_info["found_heading"],
                        "block_range": f"{tc_info['start_bid']}-{tc_info['end_bid']}",
                        "h3_children": h3_tree,
                        "status": tc_status,
                        "orphan": tc_info.get("orphan", False)
                    }

                heading_tree["h2_test_cases"] = tc_tree
                heading_tree["total_test_cases"] = len(tc_keys)


        # 6. Section 12 Header Validation
        if key == "sec12":
            sec_slice = split_map.slice(key)
            table_blocks = [b for b in sec_slice if b.get("type") == "table"]
            if table_blocks:
                first_table = table_blocks[0]
                rows = first_table.get("rows", [])
                if rows:
                    raw_headers = [str(c).strip() for c in rows[0]]
                    norm_headers = [_norm_alpha(h) for h in raw_headers]
                    
                    # Expected concepts (after alpha-normalization):
                    # 1. testcaseno / testcasenumber
                    # 2. testcasename
                    # 3. result
                    # 4. remarks
                    expected_concepts = [
                        ["testcaseno", "testcasenumber"],
                        ["testcasename"],
                        ["result"],
                        ["remarks", "remark"]
                    ]
                    
                    if len(norm_headers) < 4:
                        full_name = f"{num}. {name}"
                        issues.append(_make_issue(
                            "INVALID_SEC12_HEADERS", "HIGH",
                            f"Section 12 table expected at least 4 columns, found {len(norm_headers)}.",
                            "Add all mandatory columns to the results table, including 'Test Case No', 'Test Case Name', 'Result', and 'Remarks'.",
                            where=full_name,
                            what="The Section 12 results table is missing one or more required columns.",
                            redirect_text=full_name
                        ))
                    else:
                        for idx, concepts in enumerate(expected_concepts):
                            found_h_norm = norm_headers[idx]
                            found_h_raw  = raw_headers[idx]
                            if not any(c == found_h_norm for c in concepts):
                                full_name = f"{num}. {name}"
                                issues.append(_make_issue(
                                    "INVALID_SEC12_HEADERS", "HIGH",
                                    f"Column {idx+1} header '{found_h_raw}' does not match expected concept.",
                                    f"Update the table header '{found_h_raw}' to '{concepts[0]}'.",
                                    where=full_name,
                                    what=f"The column header '{found_h_raw}' in the Section 12 results table does not match the required header name.",
                                    redirect_text=full_name
                                ))

        # Derive status (Master Design alignment)
        # 1. If splitter already flagged it, prioritize that
        # Treat BOUNDARY_BROKEN exactly like FAIL — it must NOT allow extraction
        if status_from_splitter in ("MISSING", "BOUNDARY_BROKEN", "AMBIGUOUS", "INCOMPLETE"):
            status = status_from_splitter  # already propagated
            if status_from_splitter == "BOUNDARY_BROKEN":
                # Explicitly record it will be BLOCKED downstream
                issues.append(_make_issue(
                    "BOUNDARY_BROKEN", "HIGH",
                    f"Section {num}. {name} has a broken boundary because the next "
                    f"section is missing. Its blocks absorb the missing section's content.",
                    f"Add a Heading 1 for the missing next section to restore the boundary."
                ))
        else:
            # 2. Check for other validation issues (H3 missing etc)
            high = [i for i in issues if i["severity"] == "HIGH"]
            warn = [i for i in issues if i["severity"] in ("WARN", "MEDIUM")]
            if high:
                status = "FAIL"
            elif warn:
                status = "WARN"
            else:
                status = "PASS"

        sections[key] = {
            "status": status,
            "expected": f"{num}. {name}",
            "found": found_heading,
            "block_range": split_map.map[key]["range"] if key in split_map.map else "N/A",
            "issues": issues,
            "heading_tree": heading_tree,
        }

    # Summary
    all_statuses = []
    for s_info in sections.values():
        all_statuses.append(s_info["status"])
        tree = s_info.get("heading_tree", {})
        # Check nested test cases (Section 11)
        if "h2_test_cases" in tree:
            all_statuses.extend([tc["status"] for tc in tree["h2_test_cases"].values()])
        # Check nested subsections (Section 8)
        if "h2_children" in tree:
            all_statuses.extend([child["status"] for child in tree["h2_children"].values()])

    pass_c = sum(1 for s in all_statuses if s == "PASS")
    warn_c = sum(1 for s in all_statuses if s == "WARN")
    fail_c = sum(1 for s in all_statuses if s == "FAIL")

    if fail_c == 0 and warn_c == 0:
        overall = "PASS"
    elif fail_c == 0:
        overall = "WARN"
    elif pass_c > 0:
        overall = "PARTIAL"
    else:
        overall = "FAIL"

    return {
        "summary": {
            "status": overall,
            "total": len(EXPECTED_H1),
            "pass": pass_c,
            "warn": warn_c,
            "fail": fail_c,
        },
        "sections": sections,
    }




# ---------------------------------------------------------------------------
# PHASE 2 — CONTENT EXTRACTION (independent per section)
# ---------------------------------------------------------------------------
# Inline minimal extractors so this file works standalone
# (same logic as AI_structured_extract.py; avoids import dependency).

def _extract_simple(content: List[Dict]) -> Dict:
    """Generic: collect text paragraphs."""
    return {
        "content": [
            (it.get("text") or "").strip()
            for it in content
            if it.get("type") == "paragraph" and (it.get("text") or "").strip()
        ]
    }


def _extract_table_section(content: List[Dict]) -> Dict:
    tables, paragraphs = [], []
    for it in content:
        t = it.get("type")
        if t == "paragraph":
            txt = (it.get("text") or "").strip()
            if txt:
                paragraphs.append(txt)
        elif t == "table":
            rows = it.get("rows", [])
            if rows:
                tables.append({"headers": rows[0], "rows": rows[1:]})
    return {"paragraphs": paragraphs, "tables": tables}


def _collect_scenario_blocks(content: List[Dict]) -> List[Dict]:
    """Group bold-scenario-header blocks into scenarios."""
    scenarios, current, current_desc = [], None, ""

    def flush():
        nonlocal current, current_desc
        if current:
            scenarios.append({
                "test_scenario": current,
                "description": current_desc.strip(),
            })
        current, current_desc = None, ""

    for it in content:
        if it.get("type") != "paragraph":
            continue
        bold = (it.get("bold_formatting") or "").strip()
        text = (it.get("text") or "").strip()
        if not text:
            continue
        if _is_bold_scenario_header(text, bold):
            flush()
            current = bold or text
            rest = text[len(bold):].lstrip(": -") if text.lower().startswith(bold.lower()) else ""
            current_desc = rest + " "
        elif current:
            current_desc += text + " "
    flush()
    return scenarios


def _extract_section11(content: List[Dict]) -> Dict:
    """
    Section 11 extraction:
    H2 → test case group,  H3 → subsection,  paragraphs → content.
    """
    test_cases, current_tc, current_sub = [], None, None
    exec_order = evidence_order = 0

    H3_MAP = {
        "testcasename": "name",
        "testcasedescription": "description",
        "executionsteps": "execution",
        "testobservations": "observation",
        "testobservation": "observation",
        "evidenceprovided": "evidence",
    }

    def new_tc(heading: str) -> Dict:
        return {
            "heading": heading,
            "name": "",
            "description": "",
            "execution": [],
            "observation": "",
            "evidence": [],
        }

    def flush_tc():
        nonlocal current_tc
        if current_tc:
            current_tc["name"] = current_tc["name"].strip()
            current_tc["description"] = current_tc["description"].strip()
            current_tc["observation"] = current_tc["observation"].strip()
            test_cases.append(current_tc)

    for it in content:
        style = it.get("style", "")
        text = (it.get("text") or "").strip()
        itype = it.get("type", "")

        if itype == "paragraph" and style == "Heading 2":
            flush_tc()
            current_tc = new_tc(text)
            current_sub = None
            exec_order = evidence_order = 0
            continue

        if current_tc is None:
            continue

        if itype == "paragraph" and style == "Heading 3":
            current_sub = H3_MAP.get(_norm_alpha(text), None)
            continue

        if current_sub is None:
            continue

        if current_sub == "name" and itype == "paragraph" and text:
            current_tc["name"] += text + " "
        elif current_sub == "description" and itype == "paragraph" and text:
            current_tc["description"] += text + " "
        elif current_sub == "execution":
            if itype == "paragraph" and text:
                current_tc["execution"].append({"order": exec_order, "step": text})
                exec_order += 1
            elif itype == "image":
                current_tc["execution"].append({
                    "order": exec_order, "type": "image",
                    "image_path": it.get("image_path", ""),
                })
                exec_order += 1
        elif current_sub == "observation" and itype == "paragraph" and text:
            current_tc["observation"] += text + " "
        elif current_sub == "evidence":
            if itype == "paragraph" and text:
                current_tc["evidence"].append({"order": evidence_order, "evidence": text})
                evidence_order += 1
            elif itype == "image":
                current_tc["evidence"].append({
                    "order": evidence_order, "type": "image",
                    "image_path": it.get("image_path", ""),
                })
                evidence_order += 1

    flush_tc()
    return {"test_cases": test_cases, "total_test_cases": len(test_cases)}


def _extract_section12(content: List[Dict]) -> Dict:
    results = []
    for it in content:
        if it.get("type") != "table":
            continue
        rows = it.get("rows", [])
        if len(rows) < 2:
            continue
        for row in rows[1:]:
            if len(row) >= 4:
                tc_id = re.sub(r"[:;\s]+$", "", str(row[1] or "").strip())
                if tc_id:
                    results.append({
                        "test_case_id": tc_id,
                        "result": str(row[2] or "").strip(),
                        "remarks": str(row[3] or "").strip(),
                    })
    return {"test_results": results, "total": len(results)}


def _content_for_section(section_key: str, split_map: SplitMap) -> List[Dict]:
    """
    Convert the raw block slice into content items the extractors expect.
    Skips the first block (the heading itself).
    """
    raw = split_map.slice(section_key)
    content = []
    for b in raw[1:]:  # skip the heading block
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


SECTION_EXTRACTOR_MAP = {
    "sec1": lambda c: {
        "itsar_section_details": [
            (it.get("text") or "").strip()
            for it in c
            if it.get("type") == "paragraph" and (it.get("text") or "").strip()
        ]
    },
    "sec2": lambda c: {
        "security_requirement": [
            (it.get("text") or "").strip()
            for it in c
            if it.get("type") == "paragraph" and (it.get("text") or "").strip()
        ]
    },
    "sec3": lambda c: {
        "requirement_description": [
            (it.get("text") or "").strip()
            for it in c
            if it.get("type") == "paragraph" and (it.get("text") or "").strip()
        ]
    },
    "sec4": lambda c: _extract_table_section(c),
    "sec5": lambda c: _extract_table_section(c),
    "sec6": lambda c: {
        "preconditions": [
            {"order": i, "precondition": (it.get("text") or "").strip()}
            for i, it in enumerate(c)
            if it.get("type") == "paragraph" and (it.get("text") or "").strip()
        ]
    },
    "sec7": lambda c: {
        "test_objective": " ".join(
            (it.get("text") or "").strip()
            for it in c
            if it.get("type") == "paragraph"
        ).strip()
    },
    "sec8": lambda c: _extract_simple(c),
    "sec8_1": lambda c: {
        "test_scenarios": _collect_scenario_blocks(c),
        "total": len(_collect_scenario_blocks(c)),
    },
    "sec8_2": lambda c: _extract_simple(c),
    "sec8_3": lambda c: {
        "tools": [
            {"tool": re.sub(r'^[\u2022\u2023\u25E6\*\-]+\s*', "", (it.get("text") or "").strip())}
            for it in c
            if it.get("type") == "paragraph" and (it.get("text") or "").strip()
        ]
    },
    "sec8_4": lambda c: {
        "execution_steps": _collect_scenario_blocks(c),
        "total": len(_collect_scenario_blocks(c)),
    },
    "sec9": lambda c: {
        "expected_results": _collect_scenario_blocks(c),
        "total": len(_collect_scenario_blocks(c)),
    },
    "sec10": lambda c: _extract_simple(c),
    "sec11": lambda c: _extract_section11(c),
    "sec12": lambda c: _extract_section12(c),
}



def build_mapped_sections(blocks: List[Dict], split_map: "SplitMap") -> Dict[str, List[Dict]]:
    """
    Build a dict of section_key → list-of-blocks from the split map.

    Each value contains only the blocks that belong to that section (exclusive
    of the heading block itself, so extractors receive content-only slices).

    Keys produced:
      "front_page"              — blocks before first H1
      "sec1" … "sec12"         — top-level sections (includes the H1 heading block)
      "sec8_1" … "sec8_4"      — Section 8 H2 sub-sections (includes H2 heading block)
      "tc_1", "tc_2", …        — Section 11 test-case groups (includes H2 heading block)

    The heading block is kept at index [0] of each slice so that extractors
    can inspect the title when needed; content starts at index [1].
    """
    mapped: Dict[str, List[Dict]] = {}
    for key, info in split_map.map.items():
        idx_start = info["idx_start"]
        idx_end   = info["idx_end"]
        mapped[key] = blocks[idx_start : idx_end + 1]
    return mapped


def extract_phase2(blocks, split_map):
    """Kept for backward compat only — use run_validator() instead."""
    return {"summary": {"status": "PASS", "total": 0, "extracted": 0, "skipped": 0, "failed": 0}, "sections": {}}


def run_validator(lossless_data):
    """
    Full pipeline: split -> phase1 -> gated per-section extraction -> output.

    Extraction rule:
      Phase 1 FAIL  -> extraction BLOCKED for that section (no data emitted).
      Phase 1 WARN  -> extraction proceeds with warning flag.
      Phase 1 PASS  -> full AI-structured data extracted and embedded.

    The output "sections" dict is unified: each entry carries both the
    validation result and the extracted data (or the block reason).
    """
    blocks = lossless_data.get("blocks", [])
    document_name = lossless_data.get("document", "unknown.docx")

    print(f"[DOC] Document : {document_name}")
    print(f"[DOC] Blocks   : {len(blocks)}")

    # Build split map
    print("\n[SPLIT] Building split map ...")
    split_map = SplitMap(blocks)
    print(f"   Sections mapped: {len(split_map.map)}")

    # Build mapped sections (section_key → block list) — single source of truth
    # for downstream extractors; avoids full-document re-scan and prevents bleed.
    mapped_sections = build_mapped_sections(blocks, split_map)
    for k, v in split_map.map.items():
        print(f"   {k:20s} blocks {v['range']:>14s}  | {v['heading'][:55]}")

    # Phase 1 — Heading validation (always runs for ALL sections, never stops)
    print("\n[PHASE 1] Heading structure validation ...")
    phase1 = validate_phase1(blocks, split_map)
    p1s = phase1["summary"]
    print(f"   Status: {p1s['status']}  (pass={p1s['pass']}, warn={p1s['warn']}, fail={p1s['fail']})")
    for key, sec in phase1["sections"].items():
        icon = "[PASS]" if sec["status"] == "PASS" else "[WARN]" if sec["status"] == "WARN" else "[FAIL]"
        found_text = sec.get('found') or '--'
        print(f"   {icon} {key:8s} -> {sec['status']:5s}  | {found_text[:55]}")
        for iss in sec.get("issues", []):
            print(f"        [{iss['severity']}] {iss['type']}: {iss['message'][:80]}")

    # Phase 2 — Gated per-section extraction
    print("\n[PHASE 2] Content extraction (gated per section on Phase 1 result) ...")
    unified_sections = {}

    for es in EXPECTED_H1:
        sec_key = es["key"]
        num     = es["num"]
        name    = es["name"]

        p1_sec       = phase1["sections"].get(sec_key, {})
        p1_status    = p1_sec.get("status", "FAIL")
        p1_issues    = p1_sec.get("issues", [])
        block_range  = p1_sec.get("block_range", "N/A")
        heading_tree = p1_sec.get("heading_tree", {})

        # Section missing entirely from document
        if p1_status == "MISSING":
            unified_sections[sec_key] = {
                "section": f"{num}. {name}",
                "block_range": "N/A",
                "validation": "MISSING",
                "issues": p1_issues,
                "heading_tree": heading_tree,
                "extraction": "SKIPPED",
                "extracted_data": None,
                "skip_reason": "Section not found in document — extraction skipped.",
            }
            print(f"   [SKIP]  {sec_key} -> MISSING")
            continue

        # ✅ NEW: BOUNDARY_BROKEN is now treated as FAIL — blocks extraction
        if p1_status == "BOUNDARY_BROKEN":
            unified_sections[sec_key] = {
                "section": f"{num}. {name}",
                "block_range": block_range,
                "validation": "BOUNDARY_BROKEN",
                "issues": p1_issues,
                "heading_tree": heading_tree,
                "extraction": "BLOCKED",
                "extracted_data": None,
                "skip_reason": (
                    f"Extraction blocked: boundary is broken because the next "
                    f"expected section is missing. Blocks may be contaminated."
                ),
            }
            print(f"   [BLOCK] {sec_key} -> BOUNDARY_BROKEN (extraction blocked)")
            continue

        # Section present but Phase 1 FAIL (e.g. critical structural error)
        if p1_status == "FAIL":
            unified_sections[sec_key] = {
                "section": f"{num}. {name}",
                "block_range": block_range,
                "validation": "FAIL",
                "issues": p1_issues,
                "heading_tree": heading_tree,
                "extraction": "BLOCKED",
                "extracted_data": None,
                "skip_reason": "Extraction blocked because Phase 1 validation failed for this section.",
            }
            print(f"   [BLOCK] {sec_key} -> INVALID (Phase 1 FAIL — extraction blocked)")
            continue

        # PASS or WARN -> attempt extraction
        info    = split_map.map.get(sec_key)
        content = []

        if info:
            if sec_key == "sec11":
                # sec11 needs all blocks from its test-case sub-sections
                content = []
                # First, the main sec11 block (the heading)
                for b in split_map.slice("sec11"):
                    content.append(b)
                
                # Then, all tc_ keys in order, but ONLY those that passed validation
                tc_keys = sorted([k for k in split_map.map.keys() if k.startswith("tc_")], 
                                 key=lambda x: int(x.split("_")[1]) if x.split("_")[1].isdigit() else 0)
                for tk in tc_keys:
                    if split_map.section_dict[tk]["status"] == "PASS":
                        for b in split_map.slice(tk):
                            content.append(b)
                    else:
                        print(f"   [SKIP] {tk} -> INVALID (Excluding from extraction)")
            else:
                content = _content_for_section(sec_key, split_map)

        extractor = SECTION_EXTRACTOR_MAP.get(sec_key)
        if not extractor:
            unified_sections[sec_key] = {
                "section": f"{num}. {name}", "block_range": block_range,
                "validation": p1_status, "issues": p1_issues, "heading_tree": heading_tree,
                "extraction": "NO_EXTRACTOR", "extracted_data": None,
            }
            print(f"   [SKIP]  {sec_key} -> no extractor registered")
            continue

        try:
            extracted = extractor(content)
            unified_sections[sec_key] = {
                "section": f"{num}. {name}",
                "block_range": block_range,
                "validation": p1_status,
                "issues": p1_issues,
                "heading_tree": heading_tree,
                "extraction": "OK",
                "extracted_data": extracted,
            }
            print(f"   [OK]    {sec_key} -> {p1_status} | extracted")
        except Exception as exc:
            unified_sections[sec_key] = {
                "section": f"{num}. {name}",
                "block_range": block_range,
                "validation": p1_status,
                "issues": p1_issues,
                "heading_tree": heading_tree,
                "extraction": "EXTRACTION_ERROR",
                "extracted_data": None,
                "error": str(exc),
            }
            print(f"   [ERR]   {sec_key} -> extraction error: {exc}")

    # Embed Section 8 sub-sections (8.1-8.4) as children inside sec8's entry
    if "sec8" in unified_sections and unified_sections["sec8"].get("extraction") == "OK":
        children = {}
        for sub_es in EXPECTED_SEC8_H2:
            sub_key      = sub_es["key"]
            sub_info     = split_map.map.get(sub_key)
            sub_extractor = SECTION_EXTRACTOR_MAP.get(sub_key)
            # Gate sub-section on its Phase 1 h2_children status
            sec8_h2_tree = (unified_sections["sec8"]
                            .get("heading_tree", {})
                            .get("h2_children", {})
                            .get(sub_key, {}))
            sub_p1_ok = sec8_h2_tree.get("status") in ("PASS", "BOUNDARY_BROKEN", "AMBIGUOUS")

            if not sub_p1_ok or not sub_info or not sub_extractor:
                children[sub_key] = {
                    "section": f"{sub_es['num']}. {sub_es['name']}",
                    "block_range": sub_info["range"] if sub_info else "N/A",
                    "validation": "PASS" if sub_p1_ok else "FAIL",
                    "extraction": "BLOCKED" if not sub_p1_ok else "SKIPPED",
                    "extracted_data": None,
                    "skip_reason": "Sub-section missing or Phase 1 failed.",
                }
                continue
            sub_content = _content_for_section(sub_key, split_map)
            try:
                children[sub_key] = {
                    "section": f"{sub_es['num']}. {sub_es['name']}",
                    "block_range": sub_info["range"],
                    "validation": "PASS",
                    "extraction": "OK",
                    "extracted_data": sub_extractor(sub_content),
                }
            except Exception as exc:
                children[sub_key] = {
                    "section": f"{sub_es['num']}. {sub_es['name']}",
                    "extraction": "EXTRACTION_ERROR",
                    "error": str(exc),
                    "extracted_data": None,
                }
        unified_sections["sec8"]["subsections"] = children

    # Summary
    ok_count      = sum(1 for v in unified_sections.values() if v.get("extraction") == "OK")
    blocked_count = sum(1 for v in unified_sections.values() if v.get("extraction") in ("BLOCKED", "SKIPPED"))
    error_count   = sum(1 for v in unified_sections.values() if v.get("extraction") == "EXTRACTION_ERROR")
    # Combined Overall Status (Extraction + Validation)
    p1_status = phase1.get("summary", {}).get("status", "PASS")
    if p1_status == "PASS" and ok_count == len(EXPECTED_H1):
        overall = "PASS"
    elif p1_status == "FAIL" and ok_count == 0:
        overall = "FAIL"
    else:
        overall = "PARTIAL"

    print(f"\n[SUMMARY] overall={overall}  extracted={ok_count}  "
          f"blocked/skipped={blocked_count}  errors={error_count}")

    return {
        "document":    document_name,
        "split_map":   split_map.to_display(),
        "split_map_detail": {
            k: {"range": v["range"], "heading": v["heading"], "level": v["level"]}
            for k, v in split_map.map.items()
        },
        "mapped_sections": mapped_sections,
        "phase1":  phase1,
        "summary": {
            "status":            overall,
            "total_sections":    len(EXPECTED_H1),
            "extracted":         ok_count,
            "blocked_invalid":   blocked_count,
            "extraction_errors": error_count,
        },
        "sections": unified_sections,
    }


# ---------------------------------------------------------------------------
# CLI ENTRY POINT
# ---------------------------------------------------------------------------

def main():
    args = sys.argv[1:]
    if not args:
        print("Usage:")
        print("  python document_split_validator.py <input.docx>")
        print("  python document_split_validator.py <lossless.json>")
        print("  python document_split_validator.py <input.docx> --out output.json")
        sys.exit(1)

    input_path = Path(args[0])
    out_path_arg = None
    if "--out" in args:
        idx = args.index("--out")
        if idx + 1 < len(args):
            out_path_arg = Path(args[idx + 1])

    if not input_path.exists():
        print(f"[ERR] File not found: {input_path}")
        sys.exit(1)

    if input_path.suffix.lower() == ".json":
        print(f"[LOAD] Loading lossless JSON: {input_path}")
        with open(input_path, "r", encoding="utf-8") as f:
            lossless_data = json.load(f)
    elif input_path.suffix.lower() == ".docx":
        print(f"[LOAD] Parsing DOCX: {input_path}")
        try:
            from lossless_extract import parse_docx
        except ImportError:
            print("[ERR] lossless_extract.py not found. Run from same directory.")
            sys.exit(1)
        lossless_data = parse_docx(input_path, mode="lossless")
    else:
        print(f"[ERR] Unsupported file type: {input_path.suffix}")
        sys.exit(1)

    result   = run_validator(lossless_data)
    out_path = out_path_arg or (input_path.parent / "output.json")

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    print(f"\n[DONE] Output written -> {out_path}")

    p1_status = result["phase1"]["summary"]["status"]
    overall   = result["summary"]["status"]
    if p1_status == "PASS" and overall == "PASS":
        sys.exit(0)
    elif overall == "FAIL":
        sys.exit(2)
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
