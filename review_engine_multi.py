# """
# review_engine_multi.py — Multi-call runner for the legacy Grammarly-style UI.

# What it does
# ------------
# • Calls prompts 1..7 (parameter specialists) and 9 (merged meta-synthesis + aggregator).
# • Automatically loads prompts/8.yaml as a GLOBAL PREAMBLE and prepends it to every call
#   so all specialists + aggregator inherit the same strict-but-fair senior-editor rules.
# • Accepts simplified JSON from specialists and converts each to the legacy block the UI expects.
# • Returns final JSON for the UI (scores, per_parameter, overall_rating, strengths, weaknesses,
#   suggestions, drop_off_risks, viral_quotient).
# • overall_rating is decided by 9.yaml (model-driven, holistic) — NOT computed in Python.
# """

# from __future__ import annotations

# import os
# import re
# import json
# import time
# from typing import Any, Dict, List, Optional

# from dotenv import load_dotenv
# from langchain_google_genai import ChatGoogleGenerativeAI
# from pydantic import BaseModel, conint  # NEW: for structured aggregator output

# # NEW: use the shared normalizer/sanitizer so artifacts are cleaned before UI
# from utils1 import normalize_review_payload

# # -----------------------------------------------------------------------------#
# # Env / model
# # -----------------------------------------------------------------------------#
# load_dotenv()

# def _require_api_key() -> None:
#     if not os.getenv("GOOGLE_API_KEY"):
#         raise EnvironmentError("GOOGLE_API_KEY not found in environment/.env")

# def _make_llm(temperature: float) -> ChatGoogleGenerativeAI:
#     model = os.getenv("GEMINI_MODEL") or "gemini-2.5-flash"
#     # Deterministic-ish defaults; keep kwargs minimal for Gemini
#     return ChatGoogleGenerativeAI(model=model, temperature=temperature, top_p=0.0, top_k=1)

# # -----------------------------------------------------------------------------#
# # File helpers
# # -----------------------------------------------------------------------------#
# def _strip_bom(s: str) -> str:
#     return s.lstrip("\ufeff") if s and s.startswith("\ufeff") else s

# def _read_text(path: str) -> str:
#     with open(path, "r", encoding="utf-8") as f:
#         return _strip_bom(f.read())

# def _load_prompt(prompts_dir: str, n: int) -> str:
#     """
#     Load 'n.yaml' or 'n.yml'. If the file uses the 'content: |' convention,
#     extract the content block; otherwise, return the full file text.
#     """
#     for ext in ("yaml", "yml"):
#         p = os.path.join(prompts_dir, f"{n}.{ext}")
#         if os.path.isfile(p):
#             raw = _read_text(p).strip()
#             # Try to capture 'content: |' block; if not present, use whole file
#             m = re.search(r"^\s*content\s*:\s*\|?\s*(.*)$", raw, flags=re.S | re.I)
#             txt = (m.group(1) if m else raw).replace("\r\n", "\n")
#             return txt.strip()
#     raise FileNotFoundError(f"Prompt {n} (yaml/yml) not found in {os.path.abspath(prompts_dir)}")

# def _inject(template: str, **kwargs: str) -> str:
#     out = template
#     for k, v in kwargs.items():
#         out = out.replace("{" + k + "}", v)
#     return out

# # -----------------------------------------------------------------------------#
# # JSON extraction (for specialists that return plain JSON)
# # -----------------------------------------------------------------------------#
# _BEGIN = "BEGIN_JSON"
# _END = "END_JSON"

# def _between_tokens(text: str, start: str, end: str) -> Optional[str]:
#     i = text.find(start)
#     j = text.rfind(end)
#     if i == -1 or j == -1 or j <= i:
#         return None
#     return text[i + len(start): j].strip()

# def _fenced(text: str) -> Optional[str]:
#     m = re.search(r"```json\s*(.*?)\s*```", text, flags=re.S | re.I)
#     if m: return m.group(1).strip()
#     m = re.search(r"```\s*(.*?)\s*```", text, flags=re.S)
#     if m: return m.group(1).strip()
#     return None

# def _balanced(text: str) -> Optional[str]:
#     start = text.find("{")
#     if start == -1: return None
#     depth = 0
#     for i in range(start, len(text)):
#         ch = text[i]
#         if ch == "{":
#             depth += 1
#         elif ch == "}":
#             depth -= 1
#             if depth == 0:
#                 return text[start:i+1]
#     return None

# def _parse_json(text: str) -> Dict[str, Any]:
#     """
#     Try multiple strategies to extract valid JSON. Tolerates trailing commas.
#     """
#     for candidate in filter(None, (_between_tokens(text, _BEGIN, _END),
#                                   _fenced(text), _balanced(text))):
#         try:
#             return json.loads(candidate)
#         except Exception:
#             try:
#                 fixed = re.sub(r",\s*([\]}])", r"\1", candidate)
#                 return json.loads(fixed)
#             except Exception:
#                 continue
#     raise ValueError("Could not parse JSON from model output")

# # -----------------------------------------------------------------------------#
# # Conversions → legacy block for UI
# # -----------------------------------------------------------------------------#
# def _coerce_explanation(raw: Dict[str, Any]) -> str:
#     """
#     Accept 'explanation' (string) or 'explanation_bullets' (list[str]).
#     """
#     explanation = raw.get("explanation", "")
#     if isinstance(explanation, str) and explanation.strip():
#         return explanation.strip()
#     bullets = raw.get("explanation_bullets", [])
#     if isinstance(bullets, list) and bullets:
#         parts = [str(b).strip().rstrip(".") + "." for b in bullets if str(b).strip()]
#         return " ".join(parts)[:1400]
#     return ""

# def _to_legacy_param_block(raw: Dict[str, Any]) -> Dict[str, Any]:
#     """
#     Normalize a specialist block to the legacy shape expected by the UI.
#     - Raw 'extractions' are intentionally suppressed (we only use AOIs for inline highlights).
#     - AOIs pass through untouched (4-field schema, unlimited).
#     - 'summary' (plain human note) is passed through for right-panel display.
#     """
#     score = int(raw.get("score", 0) or 0)
#     explanation = _coerce_explanation(raw)
#     weakness = str(raw.get("weakness", "") or "").strip() or "Not present"

#     # Keep single suggestion for legacy UI AND preserve full list if present
#     suggestion = str(raw.get("suggestion", "") or "").strip()
#     suggestions_list: List[str] = []
#     if not suggestion and isinstance(raw.get("suggestions"), list) and raw["suggestions"]:
#         suggestions_list = [str(s).strip() for s in raw["suggestions"] if str(s).strip()]
#         suggestion = suggestions_list[0] if suggestions_list else ""
#     elif isinstance(raw.get("suggestions"), list):
#         suggestions_list = [str(s).strip() for s in raw["suggestions"] if str(s).strip()]
#     suggestion = suggestion or "Not present"

#     # Hide raw extractions in UI—AOIs only
#     ex: List[str] = []
#     aoi = raw.get("areas_of_improvement") or []

#     block: Dict[str, Any] = {
#         "extractions": ex,  # kept in schema but always empty
#         "score": score,
#         "explanation": explanation,
#         "weakness": weakness,
#         "suggestion": suggestion,
#         "areas_of_improvement": aoi,
#         "summary": str(raw.get("summary", "") or "").strip(),
#     }
#     if suggestions_list:
#         block["suggestions_list"] = suggestions_list
#     return block

# # -----------------------------------------------------------------------------#
# # Display-name mapping for UI & aggregator
# # -----------------------------------------------------------------------------#
# DISPLAY_BY_INDEX = {
#     1: "Suspense Building",
#     2: "Language/Tone",
#     3: "Intro + Main Hook/Cliffhanger",
#     4: "Story Structure + Flow",
#     5: "Pacing",
#     6: "Mini-Hooks (30–60s)",
#     7: "Outro (Ending)",
#     # 8 = global preamble (loaded, not called)
#     9: "Overall Summary (Aggregator)",
# }

# # -----------------------------------------------------------------------------#
# # LLM invoke with retries
# # -----------------------------------------------------------------------------#
# def _invoke_with_retries(llm: ChatGoogleGenerativeAI, prompt: str, tries: int = 3, base_delay: float = 0.8):
#     last_err = None
#     for k in range(tries):
#         try:
#             return llm.invoke(prompt)
#         except Exception as e:
#             last_err = e
#             if k < tries - 1:
#                 time.sleep(base_delay * (2 ** k))
#             else:
#                 raise
#     raise last_err  # type: ignore

# # -----------------------------------------------------------------------------#
# # Aggregator structured schema (model returns this)
# # -----------------------------------------------------------------------------#
# class AggregatorAll(BaseModel):
#     overall_rating: conint(ge=1, le=10)
#     strengths: List[str]
#     weaknesses: List[str]
#     suggestions: List[str]
#     drop_off_risks: List[str]
#     viral_quotient: str

# # -----------------------------------------------------------------------------#
# # Core runner
# # -----------------------------------------------------------------------------#
# def run_review_multi(
#     script_text: str,
#     prompts_dir: str = "prompts",
#     temperature: float = 0.0,  # default locked to 0.0
#     include_commentary: bool = False,  # kept for API parity; ignored
# ) -> str:
#     """
#     Execute prompts 1..7 and 9, prepend prompts/8.yaml as a global preamble to each call,
#     convert to legacy shape, and return BEGIN_JSON ... END_JSON for the UI.
#     """
#     _require_api_key()
#     llm = _make_llm(temperature)

#     # Load global preamble (8.yaml) if present
#     try:
#         global_preamble = _load_prompt(prompts_dir, 8).strip()
#         if global_preamble:
#             global_preamble += "\n\n"
#     except FileNotFoundError:
#         global_preamble = ""

#     # 1..7 specialists
#     scores: Dict[str, int] = {}
#     per_parameter: Dict[str, Dict[str, Any]] = {}

#     for i in range(1, 7 + 1):
#         name = DISPLAY_BY_INDEX[i]
#         tmpl = _load_prompt(prompts_dir, i)
#         prompt_body = _inject(tmpl, script=script_text)
#         prompt = f"{global_preamble}{prompt_body}"

#         try:
#             resp = _invoke_with_retries(llm, prompt)
#             raw_text = getattr(resp, "content", "") or ""
#             data = _parse_json(raw_text)
#         except Exception as e:
#             short = (str(e) or "unknown").strip()
#             raise RuntimeError(f"JSON parse failed on prompt {i} ({name}). Error: {short}")

#         block = _to_legacy_param_block(data)
#         scores[name] = int(block.get("score", 0))
#         per_parameter[name] = block

#     # Build evidence for aggregator (legacy shape only)
#     evidence = {"scores": scores, "per_parameter": per_parameter}
#     evidence_json = json.dumps(evidence, ensure_ascii=False)

#     # 9 = merged meta-synthesis + aggregator (MODEL decides overall_rating)
#     tmpl9 = _load_prompt(prompts_dir, 9)
#     prompt9_body = _inject(
#         tmpl9,
#         evidence_json=evidence_json,
#         script=script_text,
#     )
#     prompt9 = f"{global_preamble}{prompt9_body}"

#     # Ask for structured output (enforces keys and types)
#     try:
#         llm_aggr = _make_llm(temperature).with_structured_output(AggregatorAll)
#         agg: AggregatorAll = llm_aggr.invoke(prompt9)
#     except Exception as e:
#         short = (str(e) or "unknown").strip()
#         raise RuntimeError(f"Aggregator failed on prompt 9. Error: {short}")

#     # Build final payload: use model-decided overall_rating
#     final_payload: Dict[str, Any] = {
#         "scores": scores,
#         "per_parameter": per_parameter,
#         "overall_rating": int(agg.overall_rating),
#         "strengths": agg.strengths,
#         "weaknesses": agg.weaknesses,
#         "suggestions": agg.suggestions,
#         "drop_off_risks": agg.drop_off_risks,
#         "viral_quotient": agg.viral_quotient,
#     }

#     # Normalize & sanitize everything before returning
#     final_payload = normalize_review_payload(final_payload)

#     return _wrap_json(final_payload)

# def _wrap_json(payload: Dict[str, Any]) -> str:
#     return f"{_BEGIN}\n{json.dumps(payload, ensure_ascii=False)}\n{_END}\n"















# # """
# # review_engine_multi.py — Multi-call runner for the legacy Grammarly-style UI.

# # What it does
# # ------------
# # • Calls prompts 1..7 (parameter specialists) and 9 (merged meta-synthesis + aggregator).
# # • Automatically loads prompts/8.yaml as a GLOBAL PREAMBLE and prepends it to every call
# #   so all specialists + aggregator inherit the same strict-but-fair senior-editor rules.
# # • Accepts simplified JSON from specialists and converts each to the legacy block the UI expects.
# # • Returns final JSON for the UI (scores, per_parameter, overall_rating, strengths, weaknesses,
# #   suggestions, drop_off_risks, viral_quotient).
# # • overall_rating is decided by 9.yaml (model-driven, holistic) — NOT computed in Python.
# # """

# # from __future__ import annotations

# # import os
# # import re
# # import json
# # import time
# # from typing import Any, Dict, List, Optional

# # from dotenv import load_dotenv
# # from langchain_google_genai import ChatGoogleGenerativeAI
# # from pydantic import BaseModel, conint  # NEW: for structured aggregator output

# # # NEW: use the shared normalizer/sanitizer so artifacts are cleaned before UI
# # from utils1 import normalize_review_payload, _sget  # <-- bring in _sget

# # # -----------------------------------------------------------------------------#
# # # Env / model
# # # -----------------------------------------------------------------------------#
# # load_dotenv()  # local .env (no-op on Streamlit Cloud)

# # def _require_api_key() -> None:
# #     """
# #     Ensure GOOGLE_API_KEY is available for the Gemini client.
# #     Priority: os.environ -> st.secrets -> error.
# #     Also write it back to os.environ so 3P libs read it uniformly.
# #     """
# #     key = _sget("GOOGLE_API_KEY")
# #     if not key:
# #         raise EnvironmentError("GOOGLE_API_KEY not found in environment or Streamlit secrets")
# #     os.environ["GOOGLE_API_KEY"] = key  # make it visible to downstream libs

# # def _make_llm(temperature: float) -> ChatGoogleGenerativeAI:
# #     """
# #     Instantiate Gemini with a model name coming from env/secrets or default.
# #     """
# #     model = _sget("GEMINI_MODEL", "gemini-2.5-flash")
# #     # Deterministic-ish defaults; keep kwargs minimal for Gemini
# #     return ChatGoogleGenerativeAI(model=model, temperature=temperature, top_p=0.0, top_k=1)

# # # -----------------------------------------------------------------------------#
# # # File helpers
# # # -----------------------------------------------------------------------------#
# # def _strip_bom(s: str) -> str:
# #     return s.lstrip("\ufeff") if s and s.startswith("\ufeff") else s

# # def _read_text(path: str) -> str:
# #     with open(path, "r", encoding="utf-8") as f:
# #         return _strip_bom(f.read())

# # def _load_prompt(prompts_dir: str, n: int) -> str:
# #     """
# #     Load 'n.yaml' or 'n.yml'. If the file uses the 'content: |' convention,
# #     extract the content block; otherwise, return the full file text.
# #     """
# #     for ext in ("yaml", "yml"):
# #         p = os.path.join(prompts_dir, f"{n}.{ext}")
# #         if os.path.isfile(p):
# #             raw = _read_text(p).strip()
# #             # Try to capture 'content: |' block; if not present, use whole file
# #             m = re.search(r"^\s*content\s*:\s*\|?\s*(.*)$", raw, flags=re.S | re.I)
# #             txt = (m.group(1) if m else raw).replace("\r\n", "\n")
# #             return txt.strip()
# #     raise FileNotFoundError(f"Prompt {n} (yaml/yml) not found in {os.path.abspath(prompts_dir)}")

# # def _inject(template: str, **kwargs: str) -> str:
# #     out = template
# #     for k, v in kwargs.items():
# #         out = out.replace("{" + k + "}", v)
# #     return out

# # # -----------------------------------------------------------------------------#
# # # JSON extraction (for specialists that return plain JSON)
# # # -----------------------------------------------------------------------------#
# # _BEGIN = "BEGIN_JSON"
# # _END = "END_JSON"

# # def _between_tokens(text: str, start: str, end: str) -> Optional[str]:
# #     i = text.find(start)
# #     j = text.rfind(end)
# #     if i == -1 or j == -1 or j <= i:
# #         return None
# #     return text[i + len(start): j].strip()

# # def _fenced(text: str) -> Optional[str]:
# #     m = re.search(r"```json\s*(.*?)\s*```", text, flags=re.S | re.I)
# #     if m: return m.group(1).strip()
# #     m = re.search(r"```\s*(.*?)\s*```", text, flags=re.S)
# #     if m: return m.group(1).strip()
# #     return None

# # def _balanced(text: str) -> Optional[str]:
# #     start = text.find("{")
# #     if start == -1: return None
# #     depth = 0
# #     for i in range(start, len(text)):
# #         ch = text[i]
# #         if ch == "{":
# #             depth += 1
# #         elif ch == "}":
# #             depth -= 1
# #             if depth == 0:
# #                 return text[start:i+1]
# #     return None

# # def _parse_json(text: str) -> Dict[str, Any]:
# #     """
# #     Try multiple strategies to extract valid JSON. Tolerates trailing commas.
# #     """
# #     for candidate in filter(None, (_between_tokens(text, _BEGIN, _END),
# #                                   _fenced(text), _balanced(text))):
# #         try:
# #             return json.loads(candidate)
# #         except Exception:
# #             try:
# #                 fixed = re.sub(r",\s*([\]}])", r"\1", candidate)
# #                 return json.loads(fixed)
# #             except Exception:
# #                 continue
# #     raise ValueError("Could not parse JSON from model output")

# # # -----------------------------------------------------------------------------#
# # # Conversions → legacy block for UI
# # # -----------------------------------------------------------------------------#
# # def _coerce_explanation(raw: Dict[str, Any]) -> str:
# #     """
# #     Accept 'explanation' (string) or 'explanation_bullets' (list[str]).
# #     """
# #     explanation = raw.get("explanation", "")
# #     if isinstance(explanation, str) and explanation.strip():
# #         return explanation.strip()
# #     bullets = raw.get("explanation_bullets", [])
# #     if isinstance(bullets, list) and bullets:
# #         parts = [str(b).strip().rstrip(".") + "." for b in bullets if str(b).strip()]
# #         return " ".join(parts)[:1400]
# #     return ""

# # def _to_legacy_param_block(raw: Dict[str, Any]) -> Dict[str, Any]:
# #     """
# #     Normalize a specialist block to the legacy shape expected by the UI.
# #     - Raw 'extractions' are intentionally suppressed (we only use AOIs for inline highlights).
# #     - AOIs pass through untouched (4-field schema, unlimited).
# #     - 'summary' (plain human note) is passed through for right-panel display.
# #     """
# #     score = int(raw.get("score", 0) or 0)
# #     explanation = _coerce_explanation(raw)
# #     weakness = str(raw.get("weakness", "") or "").strip() or "Not present"

# #     # Keep single suggestion for legacy UI AND preserve full list if present
# #     suggestion = str(raw.get("suggestion", "") or "").strip()
# #     suggestions_list: List[str] = []
# #     if not suggestion and isinstance(raw.get("suggestions"), list) and raw["suggestions"]:
# #         suggestions_list = [str(s).strip() for s in raw["suggestions"] if str(s).strip()]
# #         suggestion = suggestions_list[0] if suggestions_list else ""
# #     elif isinstance(raw.get("suggestions"), list):
# #         suggestions_list = [str(s).strip() for s in raw["suggestions"] if str(s).strip()]
# #     suggestion = suggestion or "Not present"

# #     ex: List[str] = []  # hide raw extractions
# #     aoi = raw.get("areas_of_improvement") or []

# #     block: Dict[str, Any] = {
# #         "extractions": ex,
# #         "score": score,
# #         "explanation": explanation,
# #         "weakness": weakness,
# #         "suggestion": suggestion,
# #         "areas_of_improvement": aoi,
# #         "summary": str(raw.get("summary", "") or "").strip(),
# #     }
# #     if suggestions_list:
# #         block["suggestions_list"] = suggestions_list
# #     return block

# # # -----------------------------------------------------------------------------#
# # # Display-name mapping for UI & aggregator
# # # -----------------------------------------------------------------------------#
# # DISPLAY_BY_INDEX = {
# #     1: "Suspense Building",
# #     2: "Language/Tone",
# #     3: "Intro + Main Hook/Cliffhanger",
# #     4: "Story Structure + Flow",
# #     5: "Pacing",
# #     6: "Mini-Hooks (30–60s)",
# #     7: "Outro (Ending)",
# #     # 8 = global preamble (loaded, not called)
# #     9: "Overall Summary (Aggregator)",
# # }

# # # -----------------------------------------------------------------------------#
# # # LLM invoke with retries
# # # -----------------------------------------------------------------------------#
# # def _invoke_with_retries(llm: ChatGoogleGenerativeAI, prompt: str, tries: int = 3, base_delay: float = 0.8):
# #     last_err = None
# #     for k in range(tries):
# #         try:
# #             return llm.invoke(prompt)
# #         except Exception as e:
# #             last_err = e
# #             if k < tries - 1:
# #                 time.sleep(base_delay * (2 ** k))
# #             else:
# #                 raise
# #     raise last_err  # type: ignore

# # # -----------------------------------------------------------------------------#
# # # Aggregator structured schema (model returns this)
# # # -----------------------------------------------------------------------------#
# # class AggregatorAll(BaseModel):
# #     overall_rating: conint(ge=1, le=10)
# #     strengths: List[str]
# #     weaknesses: List[str]
# #     suggestions: List[str]
# #     drop_off_risks: List[str]
# #     viral_quotient: str

# # # -----------------------------------------------------------------------------#
# # # Core runner
# # # -----------------------------------------------------------------------------#
# # def run_review_multi(
# #     script_text: str,
# #     prompts_dir: str = "prompts",
# #     temperature: float = 0.0,  # default locked to 0.0
# #     include_commentary: bool = False,  # kept for API parity; ignored
# # ) -> str:
# #     """
# #     Execute prompts 1..7 and 9, prepend prompts/8.yaml as a global preamble to each call,
# #     convert to legacy shape, and return BEGIN_JSON ... END_JSON for the UI.
# #     """
# #     _require_api_key()
# #     llm = _make_llm(temperature)

# #     # Load global preamble (8.yaml) if present
# #     try:
# #         global_preamble = _load_prompt(prompts_dir, 8).strip()
# #         if global_preamble:
# #             global_preamble += "\n\n"
# #     except FileNotFoundError:
# #         global_preamble = ""

# #     # 1..7 specialists
# #     scores: Dict[str, int] = {}
# #     per_parameter: Dict[str, Dict[str, Any]] = {}

# #     for i in range(1, 7 + 1):
# #         name = DISPLAY_BY_INDEX[i]
# #         tmpl = _load_prompt(prompts_dir, i)
# #         prompt_body = _inject(tmpl, script=script_text)
# #         prompt = f"{global_preamble}{prompt_body}"

# #         try:
# #             resp = _invoke_with_retries(llm, prompt)
# #             raw_text = getattr(resp, "content", "") or ""
# #             data = _parse_json(raw_text)
# #         except Exception as e:
# #             short = (str(e) or "unknown").strip()
# #             raise RuntimeError(f"JSON parse failed on prompt {i} ({name}). Error: {short}")

# #         block = _to_legacy_param_block(data)
# #         scores[name] = int(block.get("score", 0))
# #         per_parameter[name] = block

# #     # Build evidence for aggregator (legacy shape only)
# #     evidence = {"scores": scores, "per_parameter": per_parameter}
# #     evidence_json = json.dumps(evidence, ensure_ascii=False)

# #     # 9 = merged meta-synthesis + aggregator (MODEL decides overall_rating)
# #     tmpl9 = _load_prompt(prompts_dir, 9)
# #     prompt9_body = _inject(
# #         tmpl9,
# #         evidence_json=evidence_json,
# #         script=script_text,
# #     )
# #     prompt9 = f"{global_preamble}{prompt9_body}"

# #     # Ask for structured output (enforces keys and types)
# #     try:
# #         llm_aggr = _make_llm(temperature).with_structured_output(AggregatorAll)
# #         agg: AggregatorAll = llm_aggr.invoke(prompt9)
# #     except Exception as e:
# #         short = (str(e) or "unknown").strip()
# #         raise RuntimeError(f"Aggregator failed on prompt 9. Error: {short}")

# #     # Build final payload: use model-decided overall_rating
# #     final_payload: Dict[str, Any] = {
# #         "scores": scores,
# #         "per_parameter": per_parameter,
# #         "overall_rating": int(agg.overall_rating),
# #         "strengths": agg.strengths,
# #         "weaknesses": agg.weaknesses,
# #         "suggestions": agg.suggestions,
# #         "drop_off_risks": agg.drop_off_risks,
# #         "viral_quotient": agg.viral_quotient,
# #     }

# #     # Normalize & sanitize everything before returning
# #     final_payload = normalize_review_payload(final_payload)

# #     return _wrap_json(final_payload)

# # def _wrap_json(payload: Dict[str, Any]) -> str:
# #     return f"{_BEGIN}\n{json.dumps(payload, ensure_ascii=False)}\n{_END}\n"





# -----------------------------------------------------------------------------
# Utilities for the Viral Script Reviewer:
# - File loaders (txt / docx with tables & textboxes / pdf)
# - JSON block extraction from model output
# - Parameter order constant
# - AOI normalization helpers (simplified 4-field schema)
# - Optional DOCX report renderer
# - Sanitizer for model artifacts
# - PLUS: _sget() helper to read from env OR Streamlit secrets
# -----------------------------------------------------------------------------

from __future__ import annotations

import os
import re
import json
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional, Iterable

from dotenv import load_dotenv
load_dotenv()  # local .env (ignored on Cloud)

try:
    import streamlit as st
except Exception:
    st = None  # type: ignore

def _sget(key: str, default: Optional[str] = None, section: Optional[str] = None) -> Optional[str]:
    """
    Safely get a secret. Priority: os.environ > st.secrets > default.
    """
    v = os.getenv(key)
    if v not in (None, ""):
        return v
    if st is not None:
        try:
            if section:
                sect = st.secrets.get(section, {})
                if isinstance(sect, dict):
                    vv = sect.get(key)
                    if vv not in (None, ""):
                        return vv
            else:
                vv = st.secrets.get(key)
                if vv not in (None, ""):
                    return vv
        except Exception:
            pass
    return default

# PDF
try:
    from PyPDF2 import PdfReader
except Exception:
    PdfReader = None  # type: ignore

# DOCX reading/writing
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

BEGIN_JSON_TOKEN = "BEGIN_JSON"
END_JSON_TOKEN = "END_JSON"

PARAM_ORDER: List[str] = [
    "Suspense Building",
    "Language/Tone",
    "Intro + Main Hook/Cliffhanger",
    "Story Structure + Flow",
    "Pacing",
    "Mini-Hooks (30–60s)",
    "Outro (Ending)",
]

AOI_KEYS: List[str] = ["quote_verbatim", "issue", "fix", "why_this_helps"]

_EMOJI_RE = re.compile(
    r'[\U0001F1E0-\U0001F6FF\U0001F900-\U0001FAFF\U00002700-\U000027BF\U0001F300-\U0001F5FF]',
    flags=re.UNICODE
)

def sanitize_editor_text(s: Optional[str]) -> str:
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

def _strip_bom(s: str) -> str:
    return s.lstrip("\ufeff") if s and s.startswith("\ufeff") else s

def _normalize_text(s: str) -> str:
    if not s:
        return ""
    s = _strip_bom(s)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("\xa0", " ")
    return s.strip()

def _iter_block_items(document: Document):
    parent = document.element.body
    for child in parent.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)

def _paragraph_text_with_breaks(p: Paragraph) -> str:
    parts: List[str] = []
    for run in p.runs:
        if run.text:
            parts.append(run.text)
        for br in run._r.xpath(".//w:br"):
            parts.append("\n")
    txt = "".join(parts)
    txt = re.sub(r'\n{3,}', '\n\n', txt)
    return txt

def _text_from_paragraph(p: Paragraph) -> str:
    return _paragraph_text_with_breaks(p) or ""

def _text_from_table(tbl: Table) -> str:
    lines: List[str] = []
    for row in tbl.rows:
        row_cells: List[str] = []
        for cell in row.cells:
            cell_bits: List[str] = []
            for para in cell.paragraphs:
                cell_bits.append(_paragraph_text_with_breaks(para))
            for nt in cell._tc.iterchildren():
                if isinstance(nt, CT_Tbl):
                    t = Table(nt, cell._parent)
                    nested = _text_from_table(t)
                    if nested:
                        cell_bits.append(nested)
            row_cells.append("\n".join([cl for cl in cell_bits if cl]))
        line = "  ".join([c for c in row_cells if c])
        lines.append(line)
    return "\n".join([ln for ln in lines if ln.strip()])

def _extract_docx_in_order(docx_path: str) -> str:
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
            texts: List[str] = []
            for t in p.findall(".//w:t", ns):
                texts.append(t.text or "")
            p_texts.append("".join(texts))
        if p_texts:
            bits.append("\n".join(p_texts))
    return _normalize_text("\n\n".join(bits))

def _load_pdf_text(path: str) -> str:
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

def load_script_file(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()

    if ext == ".txt":
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return _normalize_text(f.read())

    if ext == ".docx":
        main = _extract_docx_in_order(path)
        tbx = _extract_textboxes_from_docx(path)
        if tbx and tbx not in main:
            main = (main + ("\n\n" if main else "") + tbx).strip()
        return main

    if ext == ".pdf":
        return _load_pdf_text(path)

    return ""

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
            s2 = re.sub(r",\s*([\]}])", r"\1", s)
            try:
                return json.loads(s2)
            except Exception:
                continue
    return None

def _clean_str(x: Any) -> str:
    if x is None:
        return ""
    return str(x).strip()

def _coerce_aois(block: Dict[str, Any]) -> List[Dict[str, str]]:
    raw = block.get("areas_of_improvement") or []
    if not isinstance(raw, Iterable):
        return []
    out: List[Dict[str, str]] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        q = _clean_str(item.get("quote_verbatim") or item.get("quote") or item.get("line") or "")
        issue = _clean_str(item.get("issue") or "")
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
    if not isinstance(data, dict):
        return data

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
            block["areas_of_improvement"] = _coerce_aois(block)
            for fld in ("explanation", "weakness", "suggestion", "summary"):
                if isinstance(block.get(fld), str):
                    block[fld] = sanitize_editor_text(block[fld])
            aois = block.get("areas_of_improvement") or []
            for a in aois:
                if not isinstance(a, dict):
                    continue
                for fld in ("quote_verbatim", "issue", "fix", "why_this_helps"):
                    if isinstance(a.get(fld), str):
                        a[fld] = sanitize_editor_text(a[fld])
    return data

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
    dirpath = os.path.dirname(out_path)
    if dirpath:
        os.makedirs(dirpath, exist_ok=True)

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
