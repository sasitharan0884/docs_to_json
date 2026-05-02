"""
DOCX to JSON Parser - Lossless Structural Extraction
---------------------------------------------------
This parser extracts EXACT document structure as a flat block sequence.
No guessing, no inference, no "pending" state - just pure document facts.
"""

from __future__ import annotations
import json
import sys
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Union

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

class NumberingResolver:
    """Resolve Word auto-numbering prefixes from numbering.xml."""

    _W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    def __init__(self, docx_path):
        self._docx_path = Path(docx_path)
        # { numId: abstractNumId }
        self._num_to_abstract = {}
        # { abstractNumId: { ilvl: { numFmt, lvlText, startVal } } }
        self._abstract_defs = {}
        # { (numId, ilvl): startOverride }
        self._start_overrides = {}
        # Runtime counters: { (numId, ilvl): current_count }
        self._counters = {}
        # Last ilvl seen per numId
        self._last_ilvl = {}

        self._load_numbering_xml()

    def _qn(self, tag):
        return f"{{{self._W_NS}}}{tag}"

    def _load_numbering_xml(self):
        try:
            with zipfile.ZipFile(self._docx_path, "r") as zf:
                try:
                    xml_bytes = zf.read("word/numbering.xml")
                except KeyError:
                    return
        except Exception:
            return

        try:
            root = ET.fromstring(xml_bytes)
        except ET.ParseError:
            return

        # Build abstract definitions
        for abstract_num in root.findall("w:abstractNum", {"w": self._W_NS}):
            abs_id_raw = abstract_num.get(self._qn("abstractNumId"))
            try:
                abs_id = int(abs_id_raw)
            except (TypeError, ValueError):
                continue

            levels = {}
            for lvl in abstract_num.findall("w:lvl", {"w": self._W_NS}):
                ilvl_raw = lvl.get(self._qn("ilvl"))
                try:
                    ilvl = int(ilvl_raw)
                except (TypeError, ValueError):
                    continue

                num_fmt = "decimal"
                num_fmt_elem = lvl.find("w:numFmt", {"w": self._W_NS})
                if num_fmt_elem is not None:
                    num_fmt = num_fmt_elem.get(self._qn("val"), "decimal")

                lvl_text = "%1."
                lvl_text_elem = lvl.find("w:lvlText", {"w": self._W_NS})
                if lvl_text_elem is not None:
                    lvl_text = lvl_text_elem.get(self._qn("val"), "%1.")

                start_val = 1
                start_elem = lvl.find("w:start", {"w": self._W_NS})
                if start_elem is not None:
                    try:
                        start_val = int(start_elem.get(self._qn("val"), "1"))
                    except (TypeError, ValueError):
                        start_val = 1

                levels[ilvl] = {
                    "numFmt": num_fmt,
                    "lvlText": lvl_text,
                    "startVal": start_val,
                }

            if levels:
                self._abstract_defs[abs_id] = levels

        # Build numId -> abstractNumId mapping and start overrides
        for num in root.findall("w:num", {"w": self._W_NS}):
            num_id_raw = num.get(self._qn("numId"))
            try:
                num_id = int(num_id_raw)
            except (TypeError, ValueError):
                continue

            abs_elem = num.find("w:abstractNumId", {"w": self._W_NS})
            if abs_elem is not None:
                try:
                    abs_id = int(abs_elem.get(self._qn("val")))
                    self._num_to_abstract[num_id] = abs_id
                except (TypeError, ValueError):
                    pass

            for lvl_override in num.findall("w:lvlOverride", {"w": self._W_NS}):
                ilvl_raw = lvl_override.get(self._qn("ilvl"))
                try:
                    ilvl = int(ilvl_raw)
                except (TypeError, ValueError):
                    continue

                start_override = lvl_override.find("w:startOverride", {"w": self._W_NS})
                if start_override is not None:
                    try:
                        start_val = int(start_override.get(self._qn("val"), "1"))
                        self._start_overrides[(num_id, ilvl)] = start_val
                    except (TypeError, ValueError):
                        pass

    @staticmethod
    def _to_roman(num):
        roman_map = [
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
            (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
        ]
        result = ""
        for value, symbol in roman_map:
            while num >= value:
                result += symbol
                num -= value
        return result

    @staticmethod
    def _to_alpha(num, upper=False):
        base = ord("A") if upper else ord("a")
        result = ""
        while num > 0:
            num -= 1
            result = chr(base + (num % 26)) + result
            num //= 26
        return result

    @staticmethod
    def _bullet_glyph(lvl_text):
        if not lvl_text:
            return "\u2022"
        without_placeholders = re.sub(r"%\d+", "", lvl_text)
        glyph = without_placeholders.strip()
        return glyph if glyph else "\u2022"

    def _format_number(self, count, num_fmt, lvl_text=None):
        if num_fmt == "decimal":
            return str(count)
        if num_fmt == "lowerRoman":
            return self._to_roman(count).lower()
        if num_fmt == "upperRoman":
            return self._to_roman(count)
        if num_fmt == "lowerLetter":
            return self._to_alpha(count, upper=False)
        if num_fmt == "upperLetter":
            return self._to_alpha(count, upper=True)
        if num_fmt == "bullet":
            return self._bullet_glyph(lvl_text)
        if num_fmt == "none":
            return ""
        return str(count)

    def _get_start_val(self, num_id, ilvl, level_def):
        override = self._start_overrides.get((num_id, ilvl))
        if override is not None:
            return override
        return level_def.get("startVal", 1)

    def resolve_prefix(self, num_id, ilvl):
        abs_id = self._num_to_abstract.get(num_id)
        if abs_id is None:
            return None

        levels = self._abstract_defs.get(abs_id)
        if not levels:
            return None

        level_def = levels.get(ilvl)
        if not level_def:
            return None

        # Reset deeper levels if a higher level increments
        last_ilvl = self._last_ilvl.get(num_id)
        if last_ilvl is not None and ilvl < last_ilvl:
            for key in list(self._counters.keys()):
                key_num_id, key_ilvl = key
                if key_num_id == num_id and key_ilvl > ilvl:
                    del self._counters[key]

        # Increment current level counter
        counter_key = (num_id, ilvl)
        start_val = self._get_start_val(num_id, ilvl, level_def)
        if counter_key not in self._counters:
            self._counters[counter_key] = start_val
        else:
            self._counters[counter_key] += 1

        self._last_ilvl[num_id] = ilvl

        lvl_text = level_def.get("lvlText", "%1.")
        rendered = lvl_text

        # Replace placeholders based on lvlText
        placeholders = set(re.findall(r"%(\d+)", lvl_text))
        for placeholder in placeholders:
            try:
                level_index = int(placeholder) - 1
            except ValueError:
                continue
            level_info = levels.get(level_index)
            if not level_info:
                continue

            count_key = (num_id, level_index)
            count_val = self._counters.get(count_key)
            if count_val is None:
                count_val = self._get_start_val(num_id, level_index, level_info)

            formatted = self._format_number(
                count_val,
                level_info.get("numFmt", "decimal"),
                level_info.get("lvlText", ""),
            )
            rendered = rendered.replace(f"%{placeholder}", formatted)

        return rendered


# -------------------------------------------------
# PHASE 1: LOSSLESS STRUCTURAL EXTRACTION
# -------------------------------------------------
class StructuralExtractor:
    """
    Extracts document as flat block sequence with NO semantic interpretation.
    Each block has:
    - block_id: Unique sequential ID
    - type: paragraph, table
    """

    def _get_rendered_style(self, para: Paragraph, para_style: str) -> str:
        """
        Returns a single 'rendered_style' for the paragraph:
        - If any run has a character style, use the first non-empty run's style name
        - Otherwise fall back to paragraph style
        """
        for run in para.runs:
            if not run.text or not run.text.strip():
                continue
            run_style = run.style.name if run.style is not None else None
            if run_style:
                return run_style
        return para_style

    def _get_full_text_with_hyperlinks(self, para: Paragraph) -> str:
        """
        Reconstruct paragraph text from XML nodes so display text in complex
        hyperlink/field structures is not dropped by python-docx run parsing.
        """
        text_parts = []
        for node in para._element.iter():
            if not isinstance(node.tag, str):
                continue

            if node.tag.endswith("}t") and node.text:
                text_parts.append(node.text)
            elif node.tag.endswith("}br") or node.tag.endswith("}cr"):
                text_parts.append("\n")

        return "".join(text_parts)

    def _split_para_segments_with_styles(self, para: Paragraph, para_style: str):
        """
        Split paragraph into newline-delimited segments while preserving run-based style.
        For each segment:
          - text = concatenated text from runs up to newline
          - rendered_style = first non-empty run character style in that segment, else para_style
        """
        segments = []
        cur_text = ""
        cur_style = None  # first run style seen in this segment

        def flush():
            nonlocal cur_text, cur_style
            segments.append({
                "text": cur_text,
                "rendered_style": (cur_style if cur_style else para_style)
            })
            cur_text = ""
            cur_style = None

        for run in para.runs:
            run_text = run.text or ""
            run_style = run.style.name if run.style is not None else None

            # Split this run on '\n' because python-docx represents manual line breaks as '\n'
            parts = run_text.split("\n")
            for i, part in enumerate(parts):
                # add text piece
                cur_text += part

                # if this piece has visible text, capture run style for this segment (first one wins)
                if cur_style is None and part.strip() and run_style:
                    cur_style = run_style

                # if there was a newline after this part, flush segment and start new
                if i < len(parts) - 1:
                    flush()

        # flush last segment (even if empty; you can drop empties later)
        flush()

        # Fallback to XML-derived text when python-docx run parsing misses text
        # in complex hyperlink/field structures.
        xml_text = self._get_full_text_with_hyperlinks(para)
        if xml_text:
            run_text = "\n".join(seg["text"] for seg in segments)
            if run_text != xml_text:
                xml_segments = xml_text.split("\n")
                rebuilt_segments = []
                for idx, seg_text in enumerate(xml_segments):
                    rendered_style = para_style
                    if idx < len(segments):
                        rendered_style = segments[idx]["rendered_style"]
                    rebuilt_segments.append({
                        "text": seg_text,
                        "rendered_style": rendered_style
                    })
                return rebuilt_segments

        return segments

    def _collect_ordered_paragraph_items(self, para: Paragraph, para_style: str) -> List[Dict[str, Any]]:
        """Collect text and image items in the order they appear inside a paragraph."""
        items: List[Dict[str, Any]] = []
        text_buffer: List[str] = []
        text_style: Optional[str] = None

        def flush_text() -> None:
            nonlocal text_buffer, text_style
            text = "".join(text_buffer)
            if text.strip():
                items.append({
                    "type": "text",
                    "text": text,
                    "rendered_style": text_style or para_style,
                })
            text_buffer = []
            text_style = None

        def append_text(text: str, run_style: Optional[str]) -> None:
            nonlocal text_buffer, text_style
            if text is None:
                return

            parts = text.split("\n")
            for idx, part in enumerate(parts):
                if part:
                    text_buffer.append(part)
                    if text_style is None and part.strip() and run_style:
                        text_style = run_style
                if idx < len(parts) - 1:
                    flush_text()

        def handle_run(run_elm) -> None:
            try:
                run = Run(run_elm, para)
                run_style = run.style.name if run.style is not None else None
            except Exception:
                run_style = None

            def _process_node(node):
                for child in node:
                    if not isinstance(child.tag, str):
                        continue

                    if child.tag.endswith('}t'):
                        append_text(child.text or "", run_style)
                    elif child.tag.endswith('}tab'):
                        append_text("\t", run_style)
                    elif child.tag.endswith('}br') or child.tag.endswith('}cr'):
                        flush_text()
                    elif child.tag.endswith('}AlternateContent'):
                        for alt_child in child:
                            if isinstance(alt_child.tag, str) and alt_child.tag.endswith('}Choice'):
                                _process_node(alt_child)
                    elif child.tag.endswith('}drawing') or child.tag.endswith('}pict'):
                        flush_text()
                        for img_node in child.iter():
                            if not isinstance(img_node.tag, str):
                                continue
                            rId = None
                            if img_node.tag.endswith('}blip'):
                                rId = img_node.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                            elif img_node.tag.endswith('}imagedata'):
                                rId = img_node.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                            if rId:
                                items.append({"type": "image", "rId": rId})

            _process_node(run_elm)

        for child in para._element:
            if not isinstance(child.tag, str):
                continue

            if child.tag.endswith('}r'):
                handle_run(child)
            elif child.tag.endswith('}hyperlink'):
                for run_elm in child:
                    if isinstance(run_elm.tag, str) and run_elm.tag.endswith('}r'):
                        handle_run(run_elm)

        flush_text()
        return items

    def __init__(self, document, docx_path=None):
        self.document = document
        self.blocks = []
        self.block_counter = 0
        self._numbering_resolver = None
        self.docx_path = Path(docx_path) if docx_path else None
        if self.docx_path:
            self._numbering_resolver = NumberingResolver(self.docx_path)
            
        # Setup image directory
        self.image_dir = None
        if self.docx_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            dir_name = f"{self.docx_path.stem}_{timestamp}"
            self.image_dir_relative = Path("extracted_data") / "images" / dir_name
            self.image_dir = self.docx_path.parent / self.image_dir_relative
        
    def extract(self):
        """Extract document as flat block sequence"""
        # Iterate through document body in exact order
        for element in self.document.element.body:
            if not isinstance(element.tag, str):
                continue
                
            # Paragraph block
            if element.tag.endswith('p'):
                para = Paragraph(element, self.document)
                self._extract_paragraph(para)
            
            # Table block
            elif element.tag.endswith('tbl'):
                table = Table(element, self.document)
                self._extract_table(table)
        
        return self.blocks
    
    def _get_numpr_from_paragraph(self, para):
        """
        Return (numId, ilvl) if paragraph has numbering,
        either directly or inherited from style.
        """
        # Direct paragraph numbering
        p = para._p
        if p.pPr is not None and p.pPr.numPr is not None:
            numPr = p.pPr.numPr
            if numPr.numId is not None and numPr.ilvl is not None:
                return (
                    numPr.numId.val,
                    numPr.ilvl.val
                )

        # Style-based numbering
        style = para.style
        while style is not None:
            if style._element.pPr is not None and style._element.pPr.numPr is not None:
                numPr = style._element.pPr.numPr
                if numPr.numId is not None and numPr.ilvl is not None:
                    return (
                        numPr.numId.val,
                        numPr.ilvl.val
                    )
            style = style.base_style

        return None, None

    def _extract_bold_text(self, para: Paragraph):
        """
        Extract ONLY explicitly bold run text.
        This avoids incorrectly marking entire paragraphs as bold
        when Word styles (e.g., List Paragraph) cause inherited bold.
        """
        bold_chunks = []

        for run in para.runs:
            if not run.text:
                continue

            # Explicit run-level bold
            if run.bold is True:
                bold_chunks.append(run.text)
                continue

            # Optional: sometimes python-docx exposes bold via run.font.bold
            if run.bold is None and getattr(run.font, "bold", None) is True:
                bold_chunks.append(run.text)

        if not bold_chunks:
            return None

        # IMPORTANT: keep separators so values don't merge like "adminAdmin@123"
        return "".join([t for t in bold_chunks if t.strip()])

    def _extract_runs_metadata(self, para: Paragraph) -> List[Dict[str, Any]]:
        """
        Extract run-level formatting metadata.
        """
        runs_out = []
        for run in para.runs:
            runs_out.append({
                "text": run.text,
                "bold": bool(run.bold) if run.bold is not None else None,
                "italic": bool(run.italic) if run.italic is not None else None,
                "underline": bool(run.underline) if run.underline is not None else None,
                "style": run.style.name if run.style is not None else None,
                "font_bold": getattr(run.font, "bold", None),
                "font_italic": getattr(run.font, "italic", None)
            })
        return runs_out

    def _extract_paragraph(self, para):
        """Extract paragraph as block, resolving Word numbering from definitions."""
        block_id = self.block_counter
        self.block_counter += 1
        
        # Use XML-derived text so hyperlink display text is retained.
        text = self._get_full_text_with_hyperlinks(para)
        style = para.style.name if para.style else "Normal"

        num_id, ilvl = self._get_numpr_from_paragraph(para)
        numbering_prefix = ""

        # Resolve numbering
        rendered = None
        if num_id is not None and ilvl is not None:
            if int(num_id) != 0 and self._numbering_resolver is not None:
                rendered = self._numbering_resolver.resolve_prefix(int(num_id), int(ilvl))
                if rendered:
                    numbering_prefix = rendered.strip()


        ordered_items = self._collect_ordered_paragraph_items(para, style)
        text_item_index = 0

        for item in ordered_items:
            item_type = item.get("type")

            if item_type == "text":
                seg_text = item.get("text") or ""
                if seg_text is None or len(seg_text.strip()) == 0:
                    continue

                seg_block_id = block_id if text_item_index == 0 else self.block_counter
                if text_item_index != 0:
                    self.block_counter += 1

                seg_numbering_prefix = numbering_prefix if text_item_index == 0 else ""
                prefix_was_added = False

                # Ensure numbering_prefix is prepended to the first text segment if not already present
                if seg_numbering_prefix:
                    prefix_pattern = re.escape(seg_numbering_prefix).replace("\\ ", r"\s*")
                    if not re.match(rf"^\s*{prefix_pattern}(\s+|$)", seg_text):
                        seg_text = f"{seg_numbering_prefix} {seg_text}"
                        prefix_was_added = True
                    else:
                        prefix_was_added = True

                para_block = {
                    "block_id": seg_block_id,
                    "type": "paragraph",
                    "text": seg_text,
                    "style": style,
                    "rendered_style": item.get("rendered_style") or style,
                    "is_empty": len(seg_text.strip()) == 0,
                    "numbering": {
                        "numId": int(num_id) if num_id is not None else None,
                        "ilvl": int(ilvl) if ilvl is not None else None,
                        "rendered_prefix": seg_numbering_prefix or None,
                        "applied_to_text": prefix_was_added,
                    },
                    "bold_formatting": self._extract_bold_text(para) if text_item_index == 0 else None
                }

                self.blocks.append(para_block)
                text_item_index += 1

            elif item_type == "image":
                rId = item.get("rId")
                if not rId or not self.image_dir:
                    continue

                try:
                    image_part = self.document.part.related_parts[rId]
                    image_bytes = image_part.blob
                    content_type = image_part.content_type

                    ext = image_part.partname.split('.')[-1]
                    image_filename = f"{rId}.{ext}"

                    self.image_dir.mkdir(parents=True, exist_ok=True)

                    image_file_path = self.image_dir / image_filename
                    image_file_path.write_bytes(image_bytes)

                    relative_path_str = str(self.image_dir_relative / image_filename).replace("/", "\\")

                    img_block = {
                        "block_id": self.block_counter,
                        "type": "image",
                        "image_id": rId,
                        "path": relative_path_str,
                        "contentType": content_type
                    }
                    self.block_counter += 1
                    self.blocks.append(img_block)
                except KeyError:
                    continue
    
    def _extract_table(self, table):
        """Extract table as block"""
        block_id = self.block_counter
        self.block_counter += 1
        
        rows = []
        for row in table.rows:
            rows.append([cell.text.strip() for cell in row.cells])
        
        if not rows:
            return
        
        table_block = {
            "block_id": block_id,
            "type": "table",
            "rows": rows,
            "row_count": len(rows),
            "col_count": len(rows[0]) if rows else 0
        }
        
        self.blocks.append(table_block)


class SemanticReconstructor:
    """Compatibility stub for the unused semantic mode."""

    def __init__(self, blocks):
        self.blocks = blocks

    def reconstruct(self):
        raise NotImplementedError("Semantic reconstruction is not implemented in this extractor.")


def parse_docx(docx_path, mode='lossless'):
    document = Document(docx_path)
    extractor = StructuralExtractor(document, docx_path=docx_path)
    blocks = extractor.extract()
    
    if mode == 'lossless':
        return {
            "document": docx_path.name,
            "extraction_mode": "lossless",
            "total_blocks": len(blocks),
            "blocks": blocks
        }
    
    reconstructor = SemanticReconstructor(blocks)
    result = reconstructor.reconstruct()
    return {
        "document": docx_path.name,
        "extraction_mode": "semantic",
        "total_blocks": len(blocks),
        **result
    }

def main():
    if len(sys.argv) < 2:
        print("Usage: python lossless_extract.py <input.docx> [mode]")
        print("  mode: 'lossless' (default) or 'semantic'")
        sys.exit(1)
    
    docx_path = Path(sys.argv[1])
    mode = sys.argv[2] if len(sys.argv) > 2 else 'lossless'
    
    if not docx_path.exists():
        print(f"Error: File not found: {docx_path}")
        sys.exit(1)
    
    if docx_path.suffix.lower() != '.docx':
        print(f"Error: Only .docx files are supported. Found: {docx_path.suffix}")
        sys.exit(1)

    print(f"Parsing: {docx_path}")
    print(f"Mode: {mode}")
    
    result = parse_docx(docx_path, mode=mode)
    
    suffix = 'lossless' if mode == 'lossless' else 'semantic'
    out_path = docx_path.parent / f"{docx_path.stem}_{suffix}.json"
    
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)
    
    print(f"✅ EXTRACTION COMPLETE – JSON CREATED: {out_path}")
    print(f"📊 Total blocks: {result['total_blocks']}")
    
    para_count = sum(1 for b in result['blocks'] if b['type'] == 'paragraph')
    table_count = sum(1 for b in result['blocks'] if b['type'] == 'table')
    
    print(f"   - Paragraphs: {para_count}")
    print(f"   - Tables: {table_count}")
    
    if mode == 'semantic':
        print(f"   - Sections: {len(result['semantic']['sections'])}")
        print(f"   - Test Cases: {len(result['semantic']['test_cases'])}")

if __name__ == "__main__":
    main()
