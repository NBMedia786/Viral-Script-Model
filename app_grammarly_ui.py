# app_grammarly_ui.py — Mysterious 7 (Exact DOCX-format center pane + robust AOI highlights + Recents)
# - Preserves DOCX headings/bold/italic/underline/spacing and TABLES in the center pane
# - Highlights AOIs by mapping text ranges onto DOCX runs (offset-accurate, no format loss)
# - Page scrolls naturally (no inner scroll)
# - Fallback for TXT/PDF/paste keeps pre-wrap + highlights
# - Skips heading-like AOIs and suppresses matches overlapping real DOCX headings
# - Snaps highlight start/end to WORD BOUNDARIES; cleans AOI quotes; robust fuzzy matching
# - Heals split-word starts & merges tiny punctuation gaps
# - Bridges across zero-width / soft-hyphen / NBSP separators so spans don’t end mid-word
# - HEADER PATCH: Title never clipped when sidebar collapses
# - DOCX table merged-cell de-dupe parity + heading suppression inside table cells
# - Auto-flatten tabular DOCX to paragraph-only DOCX before analysis (no UI changes)
# - Recents: browse history (title + timestamp + overall), open any run exactly as generated
# - FIX: history loader normalizes created_at (epoch/str/missing) => consistent ISO; safe sort

import os, re, glob, json, tempfile, difflib, uuid, datetime
from pathlib import Path
from typing import Dict, Any, List, Tuple, Optional

import streamlit as st
import pandas as pd

from utils1 import extract_review_json, PARAM_ORDER, load_script_file
from review_engine_multi import run_review_multi

# ---- DOCX rendering imports (already in requirements via python-docx) ----
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table

# ---------- Folders ----------
SCRIPTS_DIR = "scripts"
PROMPTS_DIR = "prompts"
OUTPUT_DIR  = "outputs"
HISTORY_DIR = os.path.join(OUTPUT_DIR, "_history")
for p in (SCRIPTS_DIR, PROMPTS_DIR, OUTPUT_DIR, HISTORY_DIR):
    Path(p).mkdir(parents=True, exist_ok=True)

# ---------- Colors ----------
PARAM_COLORS: Dict[str, str] = {
    "Suspense Building":              "#ff6b6b",
    "Language/Tone":                  "#6b8cff",
    "Intro + Main Hook/Cliffhanger":  "#ffb86b",
    "Story Structure + Flow":         "#a78bfa",
    "Pacing":                         "#f43f5e",
    "Mini-Hooks (30–60s)":            "#eab308",
    "Outro (Ending)":                 "#8b5cf6",
}

# ---------- Config ----------
STRICT_MATCH_ONLY = False  # set True to disable fuzzy sentence fallback entirely

# ---------- App config ----------
st.set_page_config(page_title="M7 — Grammarly UI", page_icon="🕵️", layout="wide")

# ---------- Header patch & Recents card CSS ----------
def render_app_title():
    st.markdown(
        '<h1 class="app-title">Mysterious 7 — Grammarly-style Reviewer</h1>',
        unsafe_allow_html=True
    )
    st.markdown("""
    <style>
    :root { --sep:#e5e7eb; }

    /* Push content below Streamlit sticky header (covers multiple versions) */
    .stApp .block-container { padding-top: 4.25rem !important; }

    /* App title */
    .app-title{
      font-weight: 700;
      font-size: 2.1rem;
      line-height: 1.3;
      margin: 0 0 1rem 0;
      padding-left: 40px !important;
      padding-top: .25rem !important;
      white-space: normal;
      word-break: break-word;
      hyphens: auto;
      overflow: visible;
      position: relative !important;
      z-index: 10 !important;
    }
    [data-testid="collapsedControl"] { z-index: 6 !important; }
    header[data-testid="stHeader"], .stAppHeader {
      background: transparent !important;
      box-shadow: none !important;
    }
    @media (min-width: 992px){
      .app-title { padding-left: 0 !important; }
    }

    /* vertical dividers between columns */
    div[data-testid="column"]:nth-of-type(1){position:relative;}
    div[data-testid="column"]:nth-of-type(1)::after{
      content:"";position:absolute;top:0;right:0;width:1px;height:100%;background:var(--sep);
    }
    div[data-testid="column"]:nth-of-type(2){position:relative;}
    div[data-testid="column"]:nth-of-type(2)::after{
      content:"";position:absolute;top:0;right:0;width:1px;height:100%;background:var(--sep);
    }

    /* parameter chips */
    .param-row{display:flex;gap:8px;overflow-x:auto;padding:0 2px 8px;}
    .param-chip{padding:6px 10px;border-radius:999px;font-size:13px;white-space:nowrap;border:1px solid transparent;cursor:pointer;}

    /* center script container — page scrolls; no inner scroll */
    .docxwrap{
      border:1px solid #eee;
      border-radius:8px;
      padding:16px 14px 18px;
    }

    /* Typography similar to Word */
    .docxwrap .h1, .docxwrap .h2, .docxwrap .h3 { font-weight:700; margin: 10px 0 6px; }
    .docxwrap .h1 { font-size: 1.3rem; border-bottom: 2px solid #000; padding-bottom: 4px; }
    .docxwrap .h2 { font-size: 1.15rem; border-bottom: 1px solid #000; padding-bottom: 3px; }
    .docxwrap .h3 { font-size: 1.05rem; }
    .docxwrap p { margin: 10px 0; line-height: 1.7; font-family: ui-serif, Georgia, "Times New Roman", serif; }

    /* inline styles */
    .docxwrap strong { font-weight: 700; }
    .docxwrap em { font-style: italic; }
    .docxwrap u { text-decoration: underline; }

    /* DOCX-like table rendering */
    .docxwrap table { border-collapse: collapse; width: 100%; margin: 12px 0; }
    .docxwrap th, .docxwrap td {
      border: 1px solid #bbb;
      padding: 8px;
      vertical-align: top;
      font-family: ui-serif, Georgia, "Times New Roman", serif;
      line-height: 1.6;
    }

    /* AOI highlight */
    .docxwrap mark{
      padding:0 2px;
      border-radius:3px;
      border:1px solid rgba(0,0,0,.12);
    }

    /* Recents: nicer card borders */
    .rec-card{
      border:1px solid #e5e7eb;
      border-radius:10px;
      padding:14px 16px;
      margin: 10px 0 16px 0;
      box-shadow: 0 1px 0 rgba(0,0,0,.02);
      background:#fff;
    }
    .rec-title{font-weight:600; margin-bottom:.25rem;}
    .rec-meta{color:#6b7280; font-size:12.5px; margin-bottom:.4rem;}
    .rec-row{display:flex; align-items:center; justify-content:space-between; gap:12px;}
    </style>
    """, unsafe_allow_html=True)

render_app_title()

# ---------- Session ----------
for key, default in [
    ("review_ready", False),
    ("script_text", ""),
    ("base_stem", ""),
    ("data", None),
    ("spans_by_param", {}),
    ("param_choice", None),
    ("source_docx_path", None),
    ("heading_ranges", []),  # exact offsets of DOCX headings for suppression
    ("flattened_docx_path", None),   # temp flattened DOCX path (if created)
    ("flatten_used", False),         # whether we flattened for this run
    ("ui_mode", "home"),             # "home" | "review" | "recents"
]:
    st.session_state.setdefault(key, default)

# ---------- Sanitizer (for editor meta, not for the script body) ----------
_EMOJI_RE = re.compile(
    r'[\U0001F1E0-\U0001F6FF\U0001F900-\U0001FAFF\U00002700-\U000027BF\U0001F300-\U0001F5FF]',
    flags=re.UNICODE
)
def _sanitize_editor_text(s: Optional[str]) -> str:
    if not s:
        return ""
    t = str(s)
    t = re.sub(r'\bDecision\s*:\s*', '', t, flags=re.I)
    t = re.sub(r'\bScore\s*[:\-]?\s*\d+(\.\d+)?\b', '', t, flags=re.I)
    t = re.sub(r'^\s*(\(?\d+[\)\.]|\-|\•)\s*', '', t, flags=re.M)
    t = re.sub(r'^\s*[-*]\s+', '• ', t, flags=re.M)
    t = _EMOJI_RE.sub('', t)
    t = re.sub(r'[ \t]+', ' ', t)
    t = re.sub(r'\n{3,}', '\n\n', t)
    return t.strip()

# ---------- DOCX traversal ----------
def _iter_docx_blocks(document: Document):
    body = document.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)

# ---------- Auto-flatten helpers ----------
def _docx_contains_tables(path: str) -> bool:
    doc = Document(path)
    for blk in _iter_docx_blocks(doc):
        if isinstance(blk, Table):
            return True
    return False

def _copy_paragraph(dest_doc: Document, src_para: Paragraph):
    """Copy a paragraph with inline bold/italic/underline."""
    p = dest_doc.add_paragraph()
    try:
        if src_para.style and src_para.style.name:
            p.style = src_para.style.name
    except Exception:
        pass
    for run in src_para.runs:
        r = p.add_run(run.text or "")
        r.bold = run.bold
        r.italic = run.italic
        r.underline = run.underline
    return p

def flatten_docx_tables_to_longtext(source_path: str) -> str:
    """
    Create a paragraph-only DOCX from a tabular DOCX:
    - Preserves order and inline styles
    - De-dupes merged cells
    - Adds a blank line between rows and after each table
    """
    src = Document(source_path)
    new = Document()

    for blk in _iter_docx_blocks(src):
        if isinstance(blk, Paragraph):
            _copy_paragraph(new, blk)
        else:  # Table
            seen_tc_ids = set()
            for row in blk.rows:
                for cell in row.cells:
                    tc_id = id(cell._tc)
                    if tc_id in seen_tc_ids:
                        continue
                    seen_tc_ids.add(tc_id)
                    for p in cell.paragraphs:
                        _copy_paragraph(new, p)
                new.add_paragraph("")  # blank line between rows
            new.add_paragraph("")      # extra blank line after table

    fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    new.save(tmp_path)
    return tmp_path

# ---------- merged-cell–aware, heading-in-table aware builder ----------
def build_docx_text_with_meta(docx_path: str) -> Tuple[str, List[Tuple[int,int]]]:
    """
    Build plain text (offset-aligned with HTML rendering) AND collect exact
    character-offset ranges for Heading paragraphs (including those inside TABLE cells)
    to suppress AOI highlights there. Skips duplicated merged cells so offsets stay stable.
    """
    doc = Document(docx_path)
    out: List[str] = []
    heading_ranges: List[Tuple[int,int]] = []
    current_offset = 0

    def _append_and_advance(s: str):
        nonlocal current_offset
        out.append(s)
        current_offset += len(s)

    seen_tc_ids: set = set()

    for blk in _iter_docx_blocks(doc):
        if isinstance(blk, Paragraph):
            para_text = "".join(run.text or "" for run in blk.runs)
            sty = (blk.style.name or "").lower() if blk.style else ""
            if sty.startswith("heading"):
                start = current_offset
                end   = start + len(para_text)
                heading_ranges.append((start, end))
            _append_and_advance(para_text)
            _append_and_advance("\n")  # paragraph separator

        else:  # Table
            for row in blk.rows:
                row_cell_tcs = []
                for cell in row.cells:
                    tc = cell._tc
                    tc_id = id(tc)
                    row_cell_tcs.append((tc_id, cell))

                for idx, (tc_id, cell) in enumerate(row_cell_tcs):
                    if tc_id in seen_tc_ids:
                        if idx != len(row_cell_tcs) - 1:
                            _append_and_advance("\t")
                        continue

                    seen_tc_ids.add(tc_id)
                    cell_text_parts: List[str] = []
                    for i, p in enumerate(cell.paragraphs):
                        t = "".join(r.text or "" for r in p.runs)
                        sty = (p.style.name or "").lower() if p.style else ""
                        if sty.startswith("heading"):
                            hs = current_offset + sum(len(x) for x in cell_text_parts)
                            he = hs + len(t)
                            heading_ranges.append((hs, he))
                        cell_text_parts.append(t)
                        if i != len(cell.paragraphs) - 1:
                            cell_text_parts.append("\n")
                    cell_text = "".join(cell_text_parts)
                    _append_and_advance(cell_text)

                    if idx != len(row_cell_tcs) - 1:
                        _append_and_advance("\t")

                _append_and_advance("\n")
            _append_and_advance("\n")

    return "".join(out), heading_ranges

def _wrap_inline(safe_text: str, run) -> str:
    out = safe_text
    if getattr(run, "underline", False):
        out = f"<u>{out}</u>"
    if getattr(run, "italic", False):
        out = f"<em>{out}</em>"
    if getattr(run, "bold", False):
        out = f"<strong>{out}</strong>"
    return out

# ---------- Invisible/bridge characters ----------
_BRIDGE_CHARS = set("\u200b\u200c\u200d\u2060\ufeff\xa0\u00ad")  # ZW*, WJ, BOM, NBSP, soft hyphen

# ---------- renderer with merged-cell parity ----------
def render_docx_html_with_highlights(docx_path: str,
                                     highlight_spans: List[Tuple[int,int,str,str]]) -> str:
    """
    Build HTML from DOCX runs and inject <mark> by global character offsets.
    Keeps a single <mark> open across multiple runs until span ends.
    Mirrors build_docx_text_with_meta() linearization.
    """
    doc = Document(docx_path)
    spans = [s for s in highlight_spans if s[0] < s[1]]
    spans.sort(key=lambda x: x[0])

    cur_span = 0
    current_offset = 0

    def esc(s: str) -> str:
        return (s or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

    def open_mark_if_needed(html_parts, mark_state, color, end):
        if not mark_state["open"]:
            html_parts.append(
                f'<mark style="background:{color}33;border:1px solid {color};border-radius:3px;padding:0 2px;">'
            )
            mark_state.update(open=True, end=end, color=color)

    def close_mark_if_open(html_parts, mark_state):
        if mark_state["open"]:
            html_parts.append('</mark>')
            mark_state.update(open=False, end=None, color=None)

    def emit_run_text(run_text: str, run, html_parts: List[str], mark_state: Dict[str, Any]):
        nonlocal cur_span, current_offset
        t = run_text or ""
        i = 0
        while i < len(t):
            next_start, next_end, color = None, None, None
            if cur_span < len(spans):
                next_start, next_end, color, _aid = spans[cur_span]

            if not mark_state["open"]:
                if cur_span >= len(spans) or current_offset + (len(t) - i) <= next_start:
                    chunk = t[i:]
                    html_parts.append(_wrap_inline(esc(chunk), run))
                    current_offset += len(chunk)
                    break
                if current_offset < next_start:
                    take = next_start - current_offset
                    chunk = t[i:i+take]
                    html_parts.append(_wrap_inline(esc(chunk), run))
                    current_offset += take; i += take
                    continue
                open_mark_if_needed(html_parts, mark_state, color, next_end)
            else:
                take = min(mark_state["end"] - current_offset, len(t) - i)
                if take > 0:
                    chunk = t[i:i+take]
                    html_parts.append(_wrap_inline(esc(chunk), run))
                    current_offset += take; i += take
                if current_offset >= mark_state["end"]:
                    close_mark_if_open(html_parts, mark_state)
                    cur_span += 1

    html: List[str] = ['<div class="docxwrap">']
    seen_tc_ids: set = set()

    for blk in _iter_docx_blocks(doc):
        if isinstance(blk, Paragraph):
            mark_state = {"open": False, "end": None, "color": None}
            sty = (blk.style.name or "").lower() if blk.style else ""
            open_tag = '<div class="h1">' if sty.startswith("heading 1") else \
                       '<div class="h2">' if sty.startswith("heading 2") else \
                       '<div class="h3">' if sty.startswith("heading 3") else "<p>"
            close_tag = "</div>" if sty.startswith("heading") else "</p>"

            html.append(open_tag)
            for run in blk.runs:
                emit_run_text(run.text or "", run, html, mark_state)
            close_mark_if_open(html, mark_state)
            html.append(close_tag)
            current_offset += 1  # '\n'

        else:  # Table
            html.append("<table>")
            for row in blk.rows:
                html.append("<tr>")
                row_cell_tcs = [(id(cell._tc), cell) for cell in row.cells]

                for idx, (tc_id, cell) in enumerate(row_cell_tcs):
                    html.append("<td>")
                    if tc_id not in seen_tc_ids:
                        seen_tc_ids.add(tc_id)
                        for p_idx, p in enumerate(cell.paragraphs):
                            mark_state = {"open": False, "end": None, "color": None}
                            html.append("<div>")
                            for run in p.runs:
                                emit_run_text(run.text or "", run, html, mark_state)
                            close_mark_if_open(html, mark_state)
                            html.append("</div>")
                            if p_idx != len(cell.paragraphs) - 1:
                                current_offset += 1
                    html.append("</td>")
                    if idx != len(row_cell_tcs) - 1:
                        current_offset += 1  # '\t'
                html.append("</tr>")
                current_offset += 1   # row '\n'
            html.append("</table>")
            current_offset += 1       # extra '\n'

    html.append("</div>")
    return "".join(html)

# ---------- Highlight search (on plain text) ----------
def _normalize_keep_len(s: str) -> str:
    trans = {
        "\u201c": '"', "\u201d": '"', "\u2018": "'", "\u2019": "'",
        "\u2013": "-", "\u2014": "-",
        "\xa0": " ",
        "\u200b": " ", "\u200c": " ", "\u200d": " ", "\u2060": " ",
        "\ufeff": " ", "\u00ad": " ",
    }
    return (s or "").translate(str.maketrans(trans))

def _tokenize(s: str) -> List[str]:
    return re.findall(r"\w+", (s or "").lower())

def _iter_sentences_with_spans(text: str) -> List[Tuple[int,int,str]]:
    spans = []
    for m in re.finditer(r'[^.!?]+[.!?]+|\Z', text, flags=re.S):
        s, e = m.start(), m.end()
        seg = text[s:e]
        if seg.strip():
            spans.append((s, e, seg))
    return spans

def _squash_ws(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()

def _clean_quote_for_match(q: str) -> str:
    if not q:
        return ""
    q = _normalize_keep_len(q).strip()
    q = re.sub(r'^[\'"“”‘’\[\(\{<…\-\–\—\s]+', '', q)
    q = re.sub(r'[\'"“”‘’\]\)\}>…\-\–\—\s]+$', '', q)
    return _squash_ws(q)

def _snap_and_bridge_to_word(text: str, start: int, end: int, max_bridge: int = 2) -> Tuple[int,int]:
    n = len(text)
    s, e = max(0, start), max(start, end)

    def _is_invisible_ws(ch: str) -> bool:
        return ch in _BRIDGE_CHARS

    while s > 0:
        prev = text[s-1]
        cur  = text[s] if s < n else ""
        if prev.isalnum() and cur.isalnum():
            s -= 1
            continue
        j = s; brid = 0
        while j < n and _is_invisible_ws(text[j]):
            brid += 1; j += 1
            if brid > max_bridge: break
        if brid and (s-1) >= 0 and text[s-1].isalnum() and (j < n and text[j].isalnum()):
            s -= 1; continue
        break

    while e < n:
        prev = text[e-1] if e > 0 else ""
        nxt  = text[e]
        if prev.isalnum() and nxt.isalnum():
            e += 1
            continue
        j = e; brid = 0
        while j < n and _is_invisible_ws(text[j]):
            brid += 1; j += 1
            if brid > max_bridge: break
        if brid and (e-1) >= 0 and text[e-1].isalnum() and (j < n and text[j].isalnum()):
            e = j + 1; continue
        break

    while e < n and text[e] in ',"”’\')]}':
        e += 1
    return s, e

def _heal_split_word_left(text: str, start: int) -> int:
    i = start
    if i <= 1 or i >= len(text): return start
    if text[i-1] != " ": return start
    j = i - 2
    while j >= 0 and text[j].isalpha():
        j -= 1
    prev_token = text[j+1:i-1]
    if len(prev_token) == 1:
        return i - 2
    return start

def _overlaps_any(s: int, e: int, ranges: List[Tuple[int,int]]) -> bool:
    for rs, re_ in ranges:
        if e > rs and s < re_:
            return True
    return False

def _fuzzy_window_span(tl: str, nl: str, start: int, w: int) -> Tuple[float, Optional[Tuple[int,int]]]:
    window = tl[start:start+w]
    sm = difflib.SequenceMatcher(a=nl, b=window)
    blocks = [b for b in sm.get_matching_blocks() if b.size > 0]
    if not blocks:
        return 0.0, None
    coverage = sum(b.size for b in blocks) / max(1, len(nl))
    first_b = min(blocks, key=lambda b: b.b)
    last_b  = max(blocks, key=lambda b: b.b + b.size)
    s = start + first_b.b
    e = start + last_b.b + last_b.size
    return coverage, (s, e)

def find_span_smart(text: str, needle: str) -> Optional[Tuple[int,int]]:
    if not text or not needle:
        return None

    t_orig = text
    t_norm = _normalize_keep_len(text)
    n_norm = _clean_quote_for_match(needle)
    if not n_norm:
        return None

    tl = t_norm.lower()
    nl = n_norm.lower()

    i = tl.find(nl)
    if i != -1:
        s, e = _snap_and_bridge_to_word(t_orig, i, i + len(nl))
        s = _heal_split_word_left(t_orig, s)
        return (s, e)

    m = re.search(re.escape(nl).replace(r"\ ", r"\s+"), tl, flags=re.IGNORECASE)
    if m:
        s, e = _snap_and_bridge_to_word(t_orig, m.start(), m.end())
        s = _heal_split_word_left(t_orig, s)
        return (s, e)

    if not STRICT_MATCH_ONLY and len(nl) >= 12:
        w = max(60, min(240, len(nl) + 80))
        best_cov, best_span = 0.0, None
        step = max(1, w // 2)
        for start in range(0, max(1, len(tl) - w + 1), step):
            cov, se = _fuzzy_window_span(tl, nl, start, w)
            if cov > best_cov:
                best_cov, best_span = cov, se
        if best_span and best_cov >= 0.65:
            s, e = _snap_and_bridge_to_word(t_orig, best_span[0], best_span[1])
            if s > 0 and t_orig[s-1:s+1].lower() in {" the", " a", " an"}:
                s -= 1
            s = _heal_split_word_left(t_orig, s)
            return (s, e)

    if not STRICT_MATCH_ONLY:
        keys = [w for w in _tokenize(nl) if len(w) >= 4][:8]
        if len(keys) >= 2:
            kset = set(keys)
            best_score, best_span = 0.0, None
            for s, e, seg in _iter_sentences_with_spans(t_norm):
                toks = set(_tokenize(seg))
                ov = len(kset & toks)
                if ov == 0: continue
                score = ov / max(2, len(kset))
                length_pen = min(1.0, 120 / max(20, e - s))
                score *= (0.6 + 0.4 * length_pen)
                if score > best_score:
                    best_score, best_span = score, (s, min(e, s + 400))
            if best_span and best_score >= 0.35:
                s, e = _snap_and_bridge_to_word(t_orig, best_span[0], best_span[1])
                s = _heal_split_word_left(t_orig, s)
                return (s, e)
    return None

def merge_overlaps(spans: List[Tuple[int,int,str,str]]) -> List[Tuple[int,int,str,str]]:
    if not spans:
        return []
    spans.sort(key=lambda x: x[0])
    out = [spans[0]]
    for s,e,c,aid in spans[1:]:
        ps,pe,pc,paid = out[-1]
        if s <= pe and pc == c and e > pe:
            out[-1] = (ps, e, pc, paid)
        else:
            out.append((s,e,c,aid))
    return out

_PUNCT_WS = set(" \t\r\n,.;:!?)('\"[]{}-—–…") | _BRIDGE_CHARS

def merge_overlaps_and_adjacent(base_text: str,
                                spans: List[Tuple[int,int,str,str]],
                                max_gap: int = 2) -> List[Tuple[int,int,str,str]]:
    if not spans: return []
    spans = sorted(spans, key=lambda x: x[0])
    out = [spans[0]]
    for s, e, c, aid in spans[1:]:
        ps, pe, pc, paid = out[-1]
        if c == pc and s <= pe:
            out[-1] = (ps, max(pe, e), pc, paid); continue
        if c == pc and s - pe <= max_gap:
            gap = base_text[max(0, pe):max(0, s)]
            if all((ch in _PUNCT_WS) for ch in gap):
                out[-1] = (ps, e, pc, paid); continue
        out.append((s, e, c, aid))
    return out

# ---------- Heading filters ----------
def _is_heading_like(q: str) -> bool:
    if not q: return True
    s = q.strip()
    if not re.search(r'[.!?]', s):
        words = re.findall(r"[A-Za-z]+", s)
        if 1 <= len(words) <= 7:
            caps = sum(1 for w in words if w and w[0].isupper())
            if caps / max(1, len(words)) >= 0.8:
                return True
        if s.lower() in {"introduction","voiceover","outro","epilogue","prologue","credits","title","hook","horrifying discovery","final decision"}:
            return True
        if len(s) <= 3: return True
    return False

def _is_heading_context(script_text: str, s: int, e: int) -> bool:
    left = script_text.rfind("\n", 0, s) + 1
    right = script_text.find("\n", e);  right = len(script_text) if right == -1 else right
    line = script_text[left:right].strip()
    if len(line) <= 70 and not re.search(r'[.!?]', line):
        words = re.findall(r"[A-Za-z]+", line)
        if 1 <= len(words) <= 8:
            caps = sum(1 for w in words if w and w[0].isupper())
            if caps / max(1, len(words)) >= 0.7:
                return True
    return False

def _tighten_to_quote(script_text: str, span: Tuple[int,int], quote: str) -> Tuple[int,int]:
    if not span or not quote:
        return span
    s, e = span
    if e <= s or s < 0 or e > len(script_text):
        return span

    window = script_text[s:e]
    win_norm = _normalize_keep_len(window).lower()
    q_norm = _clean_quote_for_match(quote).lower()
    if not q_norm:
        return span

    i = win_norm.find(q_norm)
    if i == -1:
        m = re.search(re.escape(q_norm).replace(r"\ ", r"\s+"), win_norm, flags=re.IGNORECASE)
        if not m:
            return span
        i, j = m.start(), m.end()
    else:
        j = i + len(q_norm)

    s2, e2 = s + i, s + j
    s2, e2 = _snap_and_bridge_to_word(script_text, s2, e2)
    s2 = _heal_split_word_left(script_text, s2)
    if s2 >= s and e2 <= e and e2 > s2:
        return (s2, e2)
    return span

def build_spans_by_param(script_text: str, data: dict, heading_ranges: Optional[List[Tuple[int,int]]] = None) -> Dict[str, List[Tuple[int,int,str,str]]]:
    heading_ranges = heading_ranges or []
    raw = (data or {}).get("per_parameter", {}) or {}
    per: Dict[str, Dict[str, Any]] = {k:(v or {}) for k,v in raw.items()}
    spans_map: Dict[str, List[Tuple[int,int,str,str]]] = {p: [] for p in PARAM_ORDER}
    st.session_state["aoi_match_ranges"] = {}

    for p in spans_map.keys():
        color = PARAM_COLORS.get(p, "#ffd54f")
        blk = per.get(p, {}) or {}
        aois = blk.get("areas_of_improvement") or []
        for idx, item in enumerate(aois, start=1):
            raw_q = (item or {}).get("quote_verbatim", "") or ""
            q = _sanitize_editor_text(raw_q)
            clean = _clean_quote_for_match(re.sub(r"^[•\-\d\.\)\s]+", "", q).strip())
            if not clean: continue
            if _is_heading_like(clean): continue

            pos = find_span_smart(script_text, clean)
            if not pos: continue
            pos = _tighten_to_quote(script_text, pos, raw_q)
            s, e = pos

            if heading_ranges and _overlaps_any(s, e, heading_ranges): continue
            if _is_heading_context(script_text, s, e): continue

            aid = f"{p.replace(' ','_')}-AOI-{idx}"
            spans_map[p].append((s, e, color, aid))
            st.session_state["aoi_match_ranges"][aid] = (s, e)
    return spans_map

# ---------- History (save + load + recents UI) ----------

def _save_history_snapshot(
    title: str,
    data: dict,
    script_text: str,
    source_docx_path: Optional[str],
    heading_ranges: List[Tuple[int,int]],
    spans_by_param: Dict[str, List[Tuple[int,int,str,str]]],
    aoi_match_ranges: Dict[str, Tuple[int,int]]
):
    run_id = str(uuid.uuid4())
    now = datetime.datetime.now()
    created_at_iso = now.replace(microsecond=0).isoformat()  # local ISO
    created_at_human = now.strftime("%Y-%m-%d %H:%M:%S")

    blob = {
        "run_id": run_id,
        "title": title or "untitled",
        "created_at": created_at_iso,              # always a STRING for stability
        "created_at_human": created_at_human,
        "overall_rating": (data or {}).get("overall_rating", ""),
        "scores": (data or {}).get("scores", {}),
        "data": data or {},
        "script_text": script_text or "",
        "source_docx_path": source_docx_path,      # may be None
        "heading_ranges": heading_ranges or [],
        "spans_by_param": spans_by_param or {},
        "aoi_match_ranges": aoi_match_ranges or {},
    }
    out_fp = os.path.join(HISTORY_DIR, f"{created_at_iso.replace(':','-')}__{run_id}.json")
    with open(out_fp, "w", encoding="utf-8") as f:
        json.dump(blob, f, ensure_ascii=False, indent=2)

def _load_all_history() -> List[dict]:
    """
    Load all saved runs from outputs/_history/*.json.
    Backward-compatible:
      - If created_at is numeric (epoch), convert to ISO string.
      - If created_at is missing/invalid, use file mtime.
      - Always return created_at as string so sorting never mixes str/float.
    """
    out = []
    for fp in sorted(glob.glob(os.path.join(HISTORY_DIR, "*.json"))):
        try:
            with open(fp, "r", encoding="utf-8") as f:
                j = json.load(f)
        except Exception:
            continue

        j.setdefault("_path", fp)

        ca = j.get("created_at")
        try:
            if isinstance(ca, (int, float)):
                dt = datetime.datetime.utcfromtimestamp(float(ca))
                j["created_at"] = dt.replace(microsecond=0).isoformat() + "Z"
                if not j.get("created_at_human"):
                    j["created_at_human"] = dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")
            elif isinstance(ca, str) and ca:
                # already ISO-ish; keep
                pass
            else:
                # fallback to file mtime
                mtime = os.path.getmtime(fp)
                dt = datetime.datetime.fromtimestamp(mtime)
                j["created_at"] = dt.replace(microsecond=0).isoformat() + "Z"
                if not j.get("created_at_human"):
                    j["created_at_human"] = dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            j["created_at"] = str(ca or "")

        out.append(j)

    out.sort(key=lambda r: r.get("created_at") or "", reverse=True)
    return out

def _render_recents_centerpane():
    """Center-pane browser for saved runs. No script preview shown."""
    st.subheader("📄 Recents")

    q = st.text_input("Filter by title…", "")

    cols = st.columns([1, 4])
    with cols[0]:
        if st.button("← Back"):
            st.session_state.ui_mode = "home"
            st.rerun()

    # list
    recs = _load_all_history()
    ql = q.strip().lower()
    if ql:
        recs = [r for r in recs if ql in (r.get("title","").lower())]

    if not recs:
        st.caption("No history yet."); st.stop()

    # draw cards (no snippet/intro text)
    for rec in recs:
        run_id = rec.get("run_id")
        title  = rec.get("title") or "(untitled)"
        created_h = rec.get("created_at_human", "")
        overall = rec.get("overall_rating", "")
        with st.container():
            st.markdown(f"""
            <div class="rec-card">
              <div class="rec-row">
                <div>
                  <div class="rec-title">{title}</div>
                  <div class="rec-meta">{created_h}</div>
                  <div><strong>Overall:</strong> {overall if overall != "" else "—"}/10</div>
                </div>
                <div>
            """, unsafe_allow_html=True)

            open_key = f"open_{run_id}"
            st.button("Open", key=open_key)

            st.markdown("</div></div></div>", unsafe_allow_html=True)

            if st.session_state.get(open_key):
                # Load and restore this run
                path = rec.get("_path")
                if path and os.path.exists(path):
                    try:
                        with open(path, "r", encoding="utf-8") as f:
                            jj = json.load(f)
                    except Exception:
                        st.error("Could not load this run file."); st.stop()

                    st.session_state.script_text      = jj.get("script_text","")
                    st.session_state.base_stem        = jj.get("title","untitled")
                    st.session_state.data             = jj.get("data",{})
                    st.session_state.heading_ranges   = jj.get("heading_ranges",[])
                    st.session_state.spans_by_param   = jj.get("spans_by_param",{})
                    st.session_state.param_choice     = None
                    st.session_state.source_docx_path = jj.get("source_docx_path")
                    st.session_state.review_ready     = True
                    st.session_state["aoi_match_ranges"] = jj.get("aoi_match_ranges", {})
                    st.session_state.ui_mode          = "review"
                    st.rerun()

# ---------- Sidebar ----------
with st.sidebar:
    st.subheader("Run Settings")
    st.caption("Prompts dir"); st.code(os.path.abspath(PROMPTS_DIR))

    if st.button("📁 Recents", use_container_width=True):
        st.session_state.ui_mode = "recents"
        st.rerun()

    if st.button("🆕 New review", use_container_width=True):
        # cleanup temp flattened docx if we created one
        fp = st.session_state.get("flattened_docx_path")
        if fp and os.path.exists(fp):
            try: os.remove(fp)
            except Exception: pass
        for k in ["review_ready","script_text","base_stem","data","spans_by_param","param_choice",
                  "source_docx_path","heading_ranges","flattened_docx_path","flatten_used"]:
            st.session_state[k] = (False if k=="review_ready"
                                   else "" if k in ("script_text","base_stem")
                                   else {} if k=="spans_by_param"
                                   else [] if k=="heading_ranges"
                                   else None if k in ("source_docx_path","flattened_docx_path")
                                   else False if k=="flatten_used"
                                   else None)
        st.session_state.ui_mode = "home"
        st.rerun()

# ---------- Input screen ----------
def render_home():
    st.subheader("🎬 Script Source")
    tab_pick, tab_upload, tab_paste = st.tabs(["Pick from scripts/ folder", "Upload file", "Paste raw text"])
    selected_path = None; uploaded_file = None; pasted_text = None

    with tab_pick:
        files = [f for f in glob.glob(os.path.join(SCRIPTS_DIR, "*")) if f.lower().endswith((".txt",".docx",".pdf"))]
        rel = [os.path.relpath(f, SCRIPTS_DIR) for f in files]
        choice = st.selectbox("Choose a file", options=["— Select —"] + rel, index=0)
        if choice != "— Select —":
            selected_path = os.path.join(SCRIPTS_DIR, choice)
        if st.button("🔁 Refresh list"):
            st.rerun()

    with tab_upload:
        up = st.file_uploader("Upload .pdf / .docx / .txt", type=["pdf","docx","txt"])
        if up is not None:
            suffix = os.path.splitext(up.name)[1].lower()
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(up.read()); tmp_path = tmp.name
            uploaded_file = tmp_path

    with tab_paste:
        pasted_text = st.text_area("Paste VO / narration here", height=300)

    if st.button("🚀 Run Review", type="primary", use_container_width=True):
        base_stem = "pasted_script"; source_docx_path = None; heading_ranges = []

        if selected_path:
            base_stem = os.path.splitext(os.path.basename(selected_path))[0]
            if selected_path.lower().endswith(".docx"):
                path_to_use = selected_path
                if _docx_contains_tables(path_to_use):
                    flat = flatten_docx_tables_to_longtext(path_to_use)
                    st.session_state.flattened_docx_path = flat
                    st.session_state.flatten_used = True
                    path_to_use = flat
                script_text, heading_ranges = build_docx_text_with_meta(path_to_use)
                source_docx_path = path_to_use
            else:
                script_text = load_script_file(selected_path)

        elif uploaded_file:
            base_stem = "uploaded_script"
            if uploaded_file.lower().endswith(".docx"):
                path_to_use = uploaded_file
                if _docx_contains_tables(path_to_use):
                    flat = flatten_docx_tables_to_longtext(path_to_use)
                    st.session_state.flattened_docx_path = flat
                    st.session_state.flatten_used = True
                    path_to_use = flat
                script_text, heading_ranges = build_docx_text_with_meta(path_to_use)
                source_docx_path = path_to_use
            else:
                script_text = load_script_file(uploaded_file)

        elif pasted_text and pasted_text.strip():
            script_text = pasted_text.strip()
        else:
            st.warning("Select, upload, or paste a script first."); st.stop()

        if len(script_text.strip()) < 50:
            st.error("Extracted text looks too short. Check your file extraction."); st.stop()

        with st.spinner("Running analysis…"):
            try:
                review_text = run_review_multi(script_text=script_text, prompts_dir=PROMPTS_DIR, temperature=0.0)
            finally:
                if uploaded_file and not source_docx_path:
                    try: os.remove(uploaded_file)
                    except Exception: pass

        data = extract_review_json(review_text)
        if not data:
            st.error("JSON not detected in model output."); st.stop()

        st.session_state.script_text      = script_text
        st.session_state.base_stem        = base_stem
        st.session_state.data             = data
        st.session_state.heading_ranges   = heading_ranges
        st.session_state.spans_by_param   = build_spans_by_param(script_text, data, heading_ranges)
        st.session_state.param_choice     = None
        st.session_state.source_docx_path = source_docx_path
        st.session_state.review_ready     = True
        st.session_state.ui_mode          = "review"

        # Save snapshot for Recents
        _save_history_snapshot(
            title=base_stem,
            data=data,
            script_text=script_text,
            source_docx_path=source_docx_path,
            heading_ranges=heading_ranges,
            spans_by_param=st.session_state.spans_by_param,
            aoi_match_ranges=st.session_state.get("aoi_match_ranges", {})
        )

        st.rerun()

# ---------- Results screen ----------
def render_review():
    script_text     = st.session_state.script_text
    data            = st.session_state.data
    spans_by_param  = st.session_state.spans_by_param
    scores: Dict[str,int] = (data or {}).get("scores", {}) or {}
    source_docx_path: Optional[str] = st.session_state.source_docx_path

    left, center, right = st.columns([1.1, 2.7, 1.4], gap="large")

    with left:
        st.subheader("Final score")
        ordered = [p for p in PARAM_ORDER if p in scores]
        df = pd.DataFrame({"Parameter": ordered, "Score (1–10)": [scores.get(p, "") for p in ordered]})
        st.dataframe(df, hide_index=True, use_container_width=True)
        st.markdown(f'**Overall:** {data.get("overall_rating","—")}/10')
        st.divider()

        strengths = (data or {}).get("strengths") or []
        if not strengths:
            per = (data or {}).get("per_parameter", {}) or {}
            best = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
            for name, sc in best:
                if sc >= 8 and name in per:
                    exp = _sanitize_editor_text((per[name] or {}).get("explanation", "") or "")
                    first = re.split(r"(?<=[.!?])\s+", exp.strip())[0] if exp else f"Consistently strong {name.lower()}."
                    strengths.append(f"{name}: {first}")
                if len(strengths) >= 3: break

        def _bullets(title: str, items):
            st.markdown(f"**{title}**")
            for s in (items or []):
                if isinstance(s, str) and s.strip():
                    st.write("• " + _sanitize_editor_text(s))
            if not items:
                st.write("• —")

        _bullets("Strengths", strengths)
        _bullets("Weaknesses", data.get("weaknesses"))
        _bullets("Suggestions", data.get("suggestions"))
        _bullets("Drop-off Risks", data.get("drop_off_risks"))
        st.markdown("**Viral Quotient**"); st.write(_sanitize_editor_text(data.get("viral_quotient","—")))

    with right:
        st.subheader("Parameters")
        st.markdown('<div class="param-row">', unsafe_allow_html=True)
        for p in [p for p in PARAM_ORDER if p in scores]:
            if st.button(p, key=f"chip_{p}", help="Show inline AOI highlights for this parameter"):
                st.session_state.param_choice = p
        st.markdown('</div>', unsafe_allow_html=True)

        sel = st.session_state.param_choice
        if sel:
            blk = (data.get("per_parameter", {}) or {}).get(sel, {}) or {}
            st.markdown(f"**{sel} — Score:** {scores.get(sel,'—')}/10")

            if blk.get("explanation"):
                st.markdown("**Why this score**")
                st.write(_sanitize_editor_text(blk["explanation"]))

            if blk.get("weakness") and blk["weakness"] != "Not present":
                st.markdown("**Weakness**")
                st.write(_sanitize_editor_text(blk["weakness"]))

            if blk.get("suggestion") and blk["suggestion"] != "Not present":
                st.markdown("**Suggestion**")
                st.write(_sanitize_editor_text(blk["suggestion"]))

            aoi = blk.get("areas_of_improvement") or []
            if aoi:
                st.markdown("**Areas of Improvement**")
                for i, item in enumerate(aoi, 1):
                    popover_fn = getattr(st, "popover", None)

                    aid = f"{sel.replace(' ','_')}-AOI-{i}"
                    s_e_map = st.session_state.get("aoi_match_ranges", {})
                    if aid in s_e_map:
                        s_m, e_m = s_e_map[aid]
                        matched_line = script_text[s_m:e_m]
                        line = (matched_line if len(matched_line) <= 320 else matched_line[:300] + "…")
                    else:
                        line = _sanitize_editor_text(item.get('quote_verbatim',''))

                    issue = _sanitize_editor_text(item.get('issue',''))
                    fix = _sanitize_editor_text(item.get('fix',''))
                    why = _sanitize_editor_text(item.get('why_this_helps',''))

                    label = f"Issue {i}"
                    if callable(popover_fn):
                        with popover_fn(label):
                            if line: st.markdown(f"**Line:** {line}")
                            if issue: st.markdown(f"**Issue:** {issue}")
                            if fix: st.markdown(f"**Fix:** {fix}")
                            if why: st.caption(why)
                    else:
                        with st.expander(label, expanded=False):
                            if line: st.markdown(f"**Line:** {line}")
                            if issue: st.markdown(f"**Issue:** {issue}")
                            if fix: st.markdown(f"**Fix:** {fix}")
                            if why: st.caption(why)

            if blk.get("summary"):
                st.markdown("**Summary**")
                st.write(_sanitize_editor_text(blk["summary"]))

    with center:
        st.subheader("Script with inline highlights")
        spans = st.session_state.spans_by_param.get(st.session_state.param_choice, []) if st.session_state.param_choice else []

        if st.session_state.source_docx_path and os.path.exists(st.session_state.source_docx_path):
            html_code = render_docx_html_with_highlights(
                st.session_state.source_docx_path,
                merge_overlaps_and_adjacent(script_text, spans)
            )
            st.markdown(html_code, unsafe_allow_html=True)
        else:
            from html import escape as _esc
            text = _esc(script_text)
            spans2 = [s for s in merge_overlaps_and_adjacent(script_text, spans) if s[0] < s[1]]
            spans2.sort(key=lambda x: x[0])
            cur = 0; buf: List[str] = []
            for s,e,c,_aid in spans2:
                if s > cur: buf.append(text[cur:s])
                buf.append(f'<mark style="background:{c}33;border:1px solid {c};border-radius:3px;padding:0 2px;">{text[s:e]}</mark>')
                cur = e
            if cur < len(text): buf.append(text[cur:])
            st.markdown(
                f'<div class="docxwrap"><p style="white-space:pre-wrap; line-height:1.7; font-family:ui-serif, Georgia, Times New Roman, serif;">{"".join(buf)}</p></div>',
                unsafe_allow_html=True
            )

# ---------- Router ----------
mode = st.session_state.ui_mode
if mode == "recents":
    _render_recents_centerpane()
elif mode == "review" and st.session_state.review_ready:
    render_review()
else:
    render_home()
