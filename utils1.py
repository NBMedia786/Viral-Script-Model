

from __future__ import annotations

import os
import re
import json
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional, Iterable

# PDF
try:
    from PyPDF2 import PdfReader
except Exception:  # pragma: no cover
    PdfReader = None  # type: ignore

# DOCX reading/writing
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# -----------------------------------------------------------------------------#
# Constants
# -----------------------------------------------------------------------------#

BEGIN_JSON_TOKEN = "BEGIN_JSON"
END_JSON_TOKEN = "END_JSON"

# Final display order (7 parameters)
PARAM_ORDER: List[str] = [
    "Suspense Building",
    "Language/Tone",
    "Intro + Main Hook/Cliffhanger",
    "Story Structure + Flow",
    "Pacing",
    "Mini-Hooks (30–60s)",
    "Outro (Ending)",
]

# Areas of Improvement (AOI) canonical schema (simplified, unlimited items)
AOI_KEYS: List[str] = ["quote_verbatim", "issue", "fix", "why_this_helps"]

# -----------------------------------------------------------------------------#
# Sanitizer (shared with UI intentions)
# -----------------------------------------------------------------------------#

_EMOJI_RE = re.compile(
    r'[\U0001F1E0-\U0001F6FF\U0001F900-\U0001FAFF\U00002700-\U000027BF\U0001F300-\U0001F5FF]',
    flags=re.UNICODE
)

def sanitize_editor_text(s: Optional[str]) -> str:
    """Remove Decision/Score boilerplate, leading numeric/bullet labels, emojis, and tidy whitespace."""
    if not s:
        return ""
    t = str(s)
    t = re.sub(r'\bDecision\s*:\s*', '', t, flags=re.I)
    t = re.sub(r'\bScore\s*[:\-]?\s*\d+(\.\d+)?\b', '', t, flags=re.I)
    t = re.sub(r'^\s*(\(?\d+[\)\.]|\-|\•|\*)\s*', '', t, flags=re.M)
    t = _EMOJI_RE.sub('', t)
    t = re.sub(r'[ \t]+', ' ', t)
    t = re.sub(r'\n{3,}', '\n\n', t)
    return t.strip()

# -----------------------------------------------------------------------------#
# Helpers: normalization (preserve original spacing as much as possible)
# -----------------------------------------------------------------------------#

def _strip_bom(s: str) -> str:
    return s.lstrip("\ufeff") if s and s.startswith("\ufeff") else s

def _normalize_text(s: str) -> str:
    """
    Preserve author formatting as much as we reasonably can:
    - unify newlines, strip BOM/nbspace
    - DO NOT collapse multiple spaces/newlines aggressively
    """
    if not s:
        return ""
    s = _strip_bom(s)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("\xa0", " ")  # non-breaking space
    return s.strip()

# -----------------------------------------------------------------------------#
# DOCX extraction (paragraphs + tables + text boxes)
# -----------------------------------------------------------------------------#

def _iter_block_items(document: Document):
    """Yield paragraphs and tables in document order (tables appear where seen)."""
    parent = document.element.body
    for child in parent.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)

def _paragraph_text_with_breaks(p: Paragraph) -> str:
    """
    Extract paragraph text preserving soft line breaks (w:br) between runs.
    python-docx's .text loses explicit <w:br/> breaks; we reconstruct them.
    """
    parts: List[str] = []
    for run in p.runs:
        # append text first
        if run.text:
            parts.append(run.text)
        # detect <w:br/> under this run
        for br in run._r.xpath(".//w:br"):
            parts.append("\n")
    # Collapse multiple \n from consecutive <w:br> into single newlines
    txt = "".join(parts)
    txt = re.sub(r'\n{3,}', '\n\n', txt)
    return txt

def _text_from_paragraph(p: Paragraph) -> str:
    return _paragraph_text_with_breaks(p) or ""

def _text_from_table(tbl: Table) -> str:
    """
    Flatten a table by rows; keep cell text order; add line breaks between rows.
    Also include nested tables within cells if present.
    """
    lines: List[str] = []
    for row in tbl.rows:
        row_cells: List[str] = []
        for cell in row.cells:
            # paragraphs (with soft breaks)
            cell_bits: List[str] = []
            for para in cell.paragraphs:
                cell_bits.append(_paragraph_text_with_breaks(para))
            # nested tables
            for nt in cell._tc.iterchildren():
                if isinstance(nt, CT_Tbl):
                    t = Table(nt, cell._parent)
                    nested = _text_from_table(t)
                    if nested:
                        cell_bits.append(nested)
            row_cells.append("\n".join([cl for cl in cell_bits if cl]))
        # join cells with spacing; then add a row break
        line = "  ".join([c for c in row_cells if c])
        lines.append(line)
    return "\n".join([ln for ln in lines if ln.strip()])

def _extract_docx_in_order(docx_path: str) -> str:
    """Extract paragraphs and tables in visual order with minimal normalization."""
    doc = Document(docx_path)
    chunks: List[str] = []
    for block in _iter_block_items(doc):
        if isinstance(block, Paragraph):
            t = _text_from_paragraph(block)
            if t is not None:
                chunks.append(t)
        elif isinstance(block, Table):
            t = _text_from_table(block)
            if t is not None:
                chunks.append(t)
    text = "\n".join(chunks)
    return _normalize_text(text)

def _extract_textboxes_from_docx(docx_path: str) -> str:
    """
    Grab text inside text boxes / shapes (w:txbxContent) straight from the DOCX XML.
    Preserve paragraph-level breaks inside the textbox.
    """
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml = z.read("word/document.xml")
    except Exception:
        return ""
    root = ET.fromstring(xml)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    bits: List[str] = []
    for txbx in root.findall(".//w:txbxContent", ns):
        paras = txbx.findall(".//w:p", ns)
        p_texts: List[str] = []
        for p in paras:
            # collect runs and breaks inside this textbox paragraph
            texts: List[str] = []
            for t in p.findall(".//w:t", ns):
                texts.append(t.text or "")
            # approximate: insert newline between w:p's
            p_texts.append("".join(texts))
        if p_texts:
            bits.append("\n".join(p_texts))
    return _normalize_text("\n\n".join(bits))

# -----------------------------------------------------------------------------#
# PDF extraction
# -----------------------------------------------------------------------------#

def _load_pdf_text(path: str) -> str:
    """Basic PDF text extraction via PyPDF2/PdfReader."""
    if PdfReader is None:
        return ""
    try:
        reader = PdfReader(path)
        pages_text: List[str] = []
        for page in reader.pages:
            try:
                t = page.extract_text() or ""
            except Exception:
                t = ""
            pages_text.append(t)
        return _normalize_text("\n\n".join(pages_text))
    except Exception:
        return ""

# -----------------------------------------------------------------------------#
# Public: load_script_file
# -----------------------------------------------------------------------------#

def load_script_file(path: str) -> str:
    """
    Load a script from txt/docx/pdf.
    DOCX: grabs paragraphs + tables in visual order and appends textbox content.
    PDF: PyPDF2.
    """
    ext = os.path.splitext(path)[1].lower()

    if ext == ".txt":
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return _normalize_text(f.read())

    if ext == ".docx":
        main = _extract_docx_in_order(path)
        # Always try to append textboxes/shapes after main content (if any distinct content found)
        tbx = _extract_textboxes_from_docx(path)
        if tbx and tbx not in main:
            main = (main + ("\n\n" if main else "") + tbx).strip()
        return main

    if ext == ".pdf":
        return _load_pdf_text(path)

    return ""

# -----------------------------------------------------------------------------#
# JSON extraction from model output
# -----------------------------------------------------------------------------#

def _extract_between_tokens(text: str, start_token: str, end_token: str) -> Optional[str]:
    if not text:
        return None
    i = text.find(start_token)
    j = text.rfind(end_token)
    if i == -1 or j == -1 or j <= i:
        return None
    return text[i + len(start_token) : j].strip()

def _extract_fenced_block(text: str) -> Optional[str]:
    if not text:
        return None
    m = re.search(r"```json\s*(.*?)\s*```", text, flags=re.DOTALL | re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m = re.search(r"```\s*(.*?)\s*```", text, flags=re.DOTALL)
    if m:
        return m.group(1).strip()
    return None

def _extract_balanced_json(text: str) -> Optional[str]:
    if not text:
        return None
    start = text.find("{")
    if start == -1:
        return None
    depth = 0
    for i in range(start, len(text)):
        ch = text[i]
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[start : i + 1]
    return None

def extract_review_json(model_output: str) -> Optional[Dict[str, Any]]:
    """
    Extract and parse the structured JSON from the model output.
    Priority:
      1) BETWEEN BEGIN_JSON / END_JSON
      2) First ```json fenced block
      3) First triple-fenced block
      4) Balanced-brace fallback
    """
    if not model_output:
        return None

    candidates: List[str] = []

    between = _extract_between_tokens(model_output, BEGIN_JSON_TOKEN, END_JSON_TOKEN)
    if between:
        candidates.append(between)

    fenced = _extract_fenced_block(model_output)
    if fenced:
        candidates.append(fenced)

    balanced = _extract_balanced_json(model_output)
    if balanced:
        candidates.append(balanced)

    for s in candidates:
        try:
            return json.loads(s)
        except Exception:
            # trailing comma healing
            s2 = re.sub(r",\s*([\]}])", r"\1", s)
            try:
                return json.loads(s2)
            except Exception:
                continue
    return None

# -----------------------------------------------------------------------------#
# AOI utilities (normalize to the 4-field schema used by the new UI)
# -----------------------------------------------------------------------------#

def _clean_str(x: Any) -> str:
    if x is None:
        return ""
    return str(x).strip()

def _coerce_aois(block: Dict[str, Any]) -> List[Dict[str, str]]:
    """
    Normalize a parameter's AOIs to the simplified schema:
      {quote_verbatim, issue, fix, why_this_helps}
    - Accepts any incoming shape; pulls/renames common legacy keys if needed.
    - Filters out empty entries (no quote & no issue/fix).
    """
    raw = block.get("areas_of_improvement") or []
    if not isinstance(raw, Iterable):
        return []
    out: List[Dict[str, str]] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        q = _clean_str(item.get("quote_verbatim") or item.get("quote") or item.get("line") or "")
        issue = _clean_str(item.get("issue") or "")
        # map legacy 'edit_suggestion' → 'fix'
        fix = _clean_str(item.get("fix") or item.get("edit_suggestion") or "")
        why = _clean_str(item.get("why_this_helps") or item.get("why") or "")
        if not (q or issue or fix or why):
            continue
        out.append({
            "quote_verbatim": q[:240] if q else "",
            "issue": issue,
            "fix": fix,
            "why_this_helps": why,
        })
    return out

def normalize_review_payload(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    OPTIONAL helper: normalize AOIs across all parameters to the canonical 4-field schema
    AND sanitize explanation/weakness/suggestion/summary + AOI fields for clean output.
    Returns the same dict (mutated in place) for convenience.
    """
    if not isinstance(data, dict):
        return data

    # sanitize top-level fields if present
    for k in ("strengths", "weaknesses", "suggestions", "drop_off_risks"):
        if isinstance(data.get(k), list):
            data[k] = [sanitize_editor_text(x) for x in data[k]]

    if isinstance(data.get("viral_quotient"), str):
        data["viral_quotient"] = sanitize_editor_text(data["viral_quotient"])

    per = data.get("per_parameter") or {}
    if isinstance(per, dict):
        for _, block in per.items():
            if not isinstance(block, dict):
                continue
            # normalize AOIs
            block["areas_of_improvement"] = _coerce_aois(block)
            # sanitize common fields
            for fld in ("explanation", "weakness", "suggestion", "summary"):
                if isinstance(block.get(fld), str):
                    block[fld] = sanitize_editor_text(block[fld])
            # sanitize AOIs
            aois = block.get("areas_of_improvement") or []
            for a in aois:
                if not isinstance(a, dict): 
                    continue
                for fld in ("quote_verbatim", "issue", "fix", "why_this_helps"):
                    if isinstance(a.get(fld), str):
                        a[fld] = sanitize_editor_text(a[fld])
    return data

# -----------------------------------------------------------------------------#
# Optional: DOCX renderer (AOIs-first; raw extractions hidden)
# -----------------------------------------------------------------------------#

def _add_heading(doc: Document, text: str, size: int = 14, bold: bool = True):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)

def _add_bullets(doc: Document, items: List[str]):
    items = items or []
    for it in items:
        doc.add_paragraph(f"• {it}")

def _add_param_block(doc: Document, title: str, block: Dict[str, Any]):
    _add_heading(doc, title, size=12, bold=True)

    # AOIs first (primary editorial output for fixes)
    aois = _coerce_aois(block)
    if aois:
        doc.add_paragraph("Areas of Improvement:")
        for i, a in enumerate(aois, start=1):
            q = a.get("quote_verbatim") or ""
            issue = a.get("issue") or ""
            fix = a.get("fix") or ""
            why = a.get("why_this_helps") or ""
            if q:
                doc.add_paragraph(f"  {i}. Line: {q}")
            if issue:
                doc.add_paragraph(f"     • Issue: {issue}")
            if fix:
                doc.add_paragraph(f"     • Fix: {fix}")
            if why:
                doc.add_paragraph(f"     • Why: {why}")

    # Raw legacy extractions intentionally hidden now

    sc = block.get("score", "")
    if sc != "":
        doc.add_paragraph(f"Score: {sc}/10")
    if block.get("explanation"):
        doc.add_paragraph(f"Explanation: {block['explanation']}")
    if block.get("weakness"):
        doc.add_paragraph(f"Weakness: {block['weakness']}")
    if block.get("suggestion"):
        doc.add_paragraph(f"Suggestion: {block['suggestion']}")
    if block.get("summary"):
        doc.add_paragraph(f"Summary: {block['summary']}")
    doc.add_paragraph("")

def save_review_docx_claude_style_from_json(
    data: Dict[str, Any],
    out_path: str,
    title: Optional[str] = None,
):
    # Safe dir creation for local filenames
    dirpath = os.path.dirname(out_path)
    if dirpath:
        os.makedirs(dirpath, exist_ok=True)

    # Normalize & sanitize for clean export
    normalize_review_payload(data)

    doc = Document()

    title_text = title or "Viral Script Review"
    p = doc.add_paragraph()
    run = p.add_run(title_text)
    run.bold = True
    run.font.size = Pt(16)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph("")

    _add_heading(doc, "Parameter Analysis", size=14)
    per_param: Dict[str, Any] = data.get("per_parameter", {}) or {}
    for param in PARAM_ORDER:
        block = per_param.get(param, {}) or {}
        _add_param_block(doc, param, block)

    scores: Dict[str, Any] = data.get("scores", {}) or {}
    _add_heading(doc, "Scoring Table", size=14)
    table = doc.add_table(rows=len(PARAM_ORDER) + 1, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Parameter"
    table.rows[0].cells[1].text = "Score (1–10)"
    for i, p_name in enumerate(PARAM_ORDER, start=1):
        table.rows[i].cells[0].text = p_name
        table.rows[i].cells[1].text = str(scores.get(p_name, ""))

    doc.add_paragraph("")
    overall = data.get("overall_rating", "—")
    _add_heading(doc, f"Overall Rating: {overall}/10", size=14)

    doc.add_paragraph("")
    _add_heading(doc, "Strengths", size=13)
    _add_bullets(doc, data.get("strengths") or [])
    doc.add_paragraph("")
    _add_heading(doc, "Weaknesses", size=13)
    _add_bullets(doc, data.get("weaknesses") or [])
    doc.add_paragraph("")
    _add_heading(doc, "Suggestions", size=13)
    _add_bullets(doc, data.get("suggestions") or [])
    doc.add_paragraph("")
    _add_heading(doc, "Drop-off Risks", size=13)
    _add_bullets(doc, data.get("drop_off_risks") or [])
    doc.add_paragraph("")
    _add_heading(doc, "Viral Quotient", size=13)
    vq = data.get("viral_quotient", "")
    if vq:
        doc.add_paragraph(vq)

    doc.save(out_path)
