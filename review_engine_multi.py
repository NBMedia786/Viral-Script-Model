

from __future__ import annotations

import os
import re
import json
import time
from typing import Any, Dict, List, Optional

from dotenv import load_dotenv
from langchain_google_genai import ChatGoogleGenerativeAI
from pydantic import BaseModel, conint  # NEW: for structured aggregator output

# NEW: use the shared normalizer/sanitizer so artifacts are cleaned before UI
from utils1 import normalize_review_payload

# -----------------------------------------------------------------------------#
# Env / model
# -----------------------------------------------------------------------------#
load_dotenv()

def _require_api_key() -> None:
    if not os.getenv("GOOGLE_API_KEY"):
        raise EnvironmentError("GOOGLE_API_KEY not found in environment/.env")

def _make_llm(temperature: float) -> ChatGoogleGenerativeAI:
    model = os.getenv("GEMINI_MODEL") or "gemini-2.5-flash"
    # Deterministic-ish defaults; keep kwargs minimal for Gemini
    return ChatGoogleGenerativeAI(model=model, temperature=temperature, top_p=0.0, top_k=1)

# -----------------------------------------------------------------------------#
# File helpers
# -----------------------------------------------------------------------------#
def _strip_bom(s: str) -> str:
    return s.lstrip("\ufeff") if s and s.startswith("\ufeff") else s

def _read_text(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return _strip_bom(f.read())

def _load_prompt(prompts_dir: str, n: int) -> str:
    """
    Load 'n.yaml' or 'n.yml'. If the file uses the 'content: |' convention,
    extract the content block; otherwise, return the full file text.
    """
    for ext in ("yaml", "yml"):
        p = os.path.join(prompts_dir, f"{n}.{ext}")
        if os.path.isfile(p):
            raw = _read_text(p).strip()
            # Try to capture 'content: |' block; if not present, use whole file
            m = re.search(r"^\s*content\s*:\s*\|?\s*(.*)$", raw, flags=re.S | re.I)
            txt = (m.group(1) if m else raw).replace("\r\n", "\n")
            return txt.strip()
    raise FileNotFoundError(f"Prompt {n} (yaml/yml) not found in {os.path.abspath(prompts_dir)}")

def _inject(template: str, **kwargs: str) -> str:
    out = template
    for k, v in kwargs.items():
        out = out.replace("{" + k + "}", v)
    return out

# -----------------------------------------------------------------------------#
# JSON extraction (for specialists that return plain JSON)
# -----------------------------------------------------------------------------#
_BEGIN = "BEGIN_JSON"
_END = "END_JSON"

def _between_tokens(text: str, start: str, end: str) -> Optional[str]:
    i = text.find(start)
    j = text.rfind(end)
    if i == -1 or j == -1 or j <= i:
        return None
    return text[i + len(start): j].strip()

def _fenced(text: str) -> Optional[str]:
    m = re.search(r"```json\s*(.*?)\s*```", text, flags=re.S | re.I)
    if m: return m.group(1).strip()
    m = re.search(r"```\s*(.*?)\s*```", text, flags=re.S)
    if m: return m.group(1).strip()
    return None

def _balanced(text: str) -> Optional[str]:
    start = text.find("{")
    if start == -1: return None
    depth = 0
    for i in range(start, len(text)):
        ch = text[i]
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[start:i+1]
    return None

def _parse_json(text: str) -> Dict[str, Any]:
    """
    Try multiple strategies to extract valid JSON. Tolerates trailing commas.
    """
    for candidate in filter(None, (_between_tokens(text, _BEGIN, _END),
                                  _fenced(text), _balanced(text))):
        try:
            return json.loads(candidate)
        except Exception:
            try:
                fixed = re.sub(r",\s*([\]}])", r"\1", candidate)
                return json.loads(fixed)
            except Exception:
                continue
    raise ValueError("Could not parse JSON from model output")

# -----------------------------------------------------------------------------#
# Conversions → legacy block for UI
# -----------------------------------------------------------------------------#
def _coerce_explanation(raw: Dict[str, Any]) -> str:
    """
    Accept 'explanation' (string) or 'explanation_bullets' (list[str]).
    """
    explanation = raw.get("explanation", "")
    if isinstance(explanation, str) and explanation.strip():
        return explanation.strip()
    bullets = raw.get("explanation_bullets", [])
    if isinstance(bullets, list) and bullets:
        parts = [str(b).strip().rstrip(".") + "." for b in bullets if str(b).strip()]
        return " ".join(parts)[:1400]
    return ""

def _to_legacy_param_block(raw: Dict[str, Any]) -> Dict[str, Any]:
    """
    Normalize a specialist block to the legacy shape expected by the UI.
    - Raw 'extractions' are intentionally suppressed (we only use AOIs for inline highlights).
    - AOIs pass through untouched (4-field schema, unlimited).
    - 'summary' (plain human note) is passed through for right-panel display.
    """
    score = int(raw.get("score", 0) or 0)
    explanation = _coerce_explanation(raw)
    weakness = str(raw.get("weakness", "") or "").strip() or "Not present"

    # Keep single suggestion for legacy UI AND preserve full list if present
    suggestion = str(raw.get("suggestion", "") or "").strip()
    suggestions_list: List[str] = []
    if not suggestion and isinstance(raw.get("suggestions"), list) and raw["suggestions"]:
        suggestions_list = [str(s).strip() for s in raw["suggestions"] if str(s).strip()]
        suggestion = suggestions_list[0] if suggestions_list else ""
    elif isinstance(raw.get("suggestions"), list):
        suggestions_list = [str(s).strip() for s in raw["suggestions"] if str(s).strip()]
    suggestion = suggestion or "Not present"

    # Hide raw extractions in UI—AOIs only
    ex: List[str] = []
    aoi = raw.get("areas_of_improvement") or []

    block: Dict[str, Any] = {
        "extractions": ex,  # kept in schema but always empty
        "score": score,
        "explanation": explanation,
        "weakness": weakness,
        "suggestion": suggestion,
        "areas_of_improvement": aoi,
        "summary": str(raw.get("summary", "") or "").strip(),
    }
    if suggestions_list:
        block["suggestions_list"] = suggestions_list
    return block

# -----------------------------------------------------------------------------#
# Display-name mapping for UI & aggregator
# -----------------------------------------------------------------------------#
DISPLAY_BY_INDEX = {
    1: "Suspense Building",
    2: "Language/Tone",
    3: "Intro + Main Hook/Cliffhanger",
    4: "Story Structure + Flow",
    5: "Pacing",
    6: "Mini-Hooks (30–60s)",
    7: "Outro (Ending)",
    # 8 = global preamble (loaded, not called)
    9: "Overall Summary (Aggregator)",
}

# -----------------------------------------------------------------------------#
# LLM invoke with retries
# -----------------------------------------------------------------------------#
def _invoke_with_retries(llm: ChatGoogleGenerativeAI, prompt: str, tries: int = 3, base_delay: float = 0.8):
    last_err = None
    for k in range(tries):
        try:
            return llm.invoke(prompt)
        except Exception as e:
            last_err = e
            if k < tries - 1:
                time.sleep(base_delay * (2 ** k))
            else:
                raise
    raise last_err  # type: ignore

# -----------------------------------------------------------------------------#
# Aggregator structured schema (model returns this)
# -----------------------------------------------------------------------------#
class AggregatorAll(BaseModel):
    overall_rating: conint(ge=1, le=10)
    strengths: List[str]
    weaknesses: List[str]
    suggestions: List[str]
    drop_off_risks: List[str]
    viral_quotient: str

# -----------------------------------------------------------------------------#
# Core runner
# -----------------------------------------------------------------------------#
def run_review_multi(
    script_text: str,
    prompts_dir: str = "prompts",
    temperature: float = 0.0,  # default locked to 0.0
    include_commentary: bool = False,  # kept for API parity; ignored
) -> str:
    """
    Execute prompts 1..7 and 9, prepend prompts/8.yaml as a global preamble to each call,
    convert to legacy shape, and return BEGIN_JSON ... END_JSON for the UI.
    """
    _require_api_key()
    llm = _make_llm(temperature)

    # Load global preamble (8.yaml) if present
    try:
        global_preamble = _load_prompt(prompts_dir, 8).strip()
        if global_preamble:
            global_preamble += "\n\n"
    except FileNotFoundError:
        global_preamble = ""

    # 1..7 specialists
    scores: Dict[str, int] = {}
    per_parameter: Dict[str, Dict[str, Any]] = {}

    for i in range(1, 7 + 1):
        name = DISPLAY_BY_INDEX[i]
        tmpl = _load_prompt(prompts_dir, i)
        prompt_body = _inject(tmpl, script=script_text)
        prompt = f"{global_preamble}{prompt_body}"

        try:
            resp = _invoke_with_retries(llm, prompt)
            raw_text = getattr(resp, "content", "") or ""
            data = _parse_json(raw_text)
        except Exception as e:
            short = (str(e) or "unknown").strip()
            raise RuntimeError(f"JSON parse failed on prompt {i} ({name}). Error: {short}")

        block = _to_legacy_param_block(data)
        scores[name] = int(block.get("score", 0))
        per_parameter[name] = block

    # Build evidence for aggregator (legacy shape only)
    evidence = {"scores": scores, "per_parameter": per_parameter}
    evidence_json = json.dumps(evidence, ensure_ascii=False)

    # 9 = merged meta-synthesis + aggregator (MODEL decides overall_rating)
    tmpl9 = _load_prompt(prompts_dir, 9)
    prompt9_body = _inject(
        tmpl9,
        evidence_json=evidence_json,
        script=script_text,
    )
    prompt9 = f"{global_preamble}{prompt9_body}"

    # Ask for structured output (enforces keys and types)
    try:
        llm_aggr = _make_llm(temperature).with_structured_output(AggregatorAll)
        agg: AggregatorAll = llm_aggr.invoke(prompt9)
    except Exception as e:
        short = (str(e) or "unknown").strip()
        raise RuntimeError(f"Aggregator failed on prompt 9. Error: {short}")

    # Build final payload: use model-decided overall_rating
    final_payload: Dict[str, Any] = {
        "scores": scores,
        "per_parameter": per_parameter,
        "overall_rating": int(agg.overall_rating),
        "strengths": agg.strengths,
        "weaknesses": agg.weaknesses,
        "suggestions": agg.suggestions,
        "drop_off_risks": agg.drop_off_risks,
        "viral_quotient": agg.viral_quotient,
    }

    # Normalize & sanitize everything before returning
    final_payload = normalize_review_payload(final_payload)

    return _wrap_json(final_payload)

def _wrap_json(payload: Dict[str, Any]) -> str:
    return f"{_BEGIN}\n{json.dumps(payload, ensure_ascii=False)}\n{_END}\n"
