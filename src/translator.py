"""Translate extracted French text to Korean with a local llama.cpp server."""

import re

import requests

from config import LLAMA_MODEL, LLAMA_SERVER_URL, MAX_TOKENS, REQUEST_TIMEOUT, TEMPERATURE


CHAT_URL = LLAMA_SERVER_URL.replace("/completion", "/chat/completions")

KNOWN_ACRONYMS = {
    "NAK",
    "KOICA",
    "SCA",
    "GED",
    "SGEA",
    "ADM",
    "ACI",
    "DAAF",
    "DCT",
    "SPDI",
    "SGRH",
    "CNA",
    "SGG",
    "MAD",
    "MDH",
    "KPI",
    "UI",
    "UX",
    "SAE",
    "AMS",
    "OAIS",
    "ISO",
    "RH",
    "IT",
    "R&D",
}

SKIP_SYMBOLS = {"", "----", "->", "→", "•", "-", "–", "—"}
NUMBER_ONLY_RE = re.compile(r"^[\d\s.,:%+<>()/-]+(?:MDH|MAD|%)?$", re.IGNORECASE)
EMAIL_RE = re.compile(r"^[\w.\-+]+@[\w.\-]+\.[a-zA-Z]{2,}$")
URL_RE = re.compile(r"^(https?://|www\.)\S+$")
PHONE_RE = re.compile(r"^[+\d\s\-().]{7,20}$")


def normalize_acronym(text):
    """Normalize acronym-like text before whitelist comparison."""
    return text.strip().replace(" ", "").replace("/", "").replace("-", "").upper()


def should_skip(text):
    """Return True when text should be preserved without LLM translation."""
    stripped = text.strip()
    if stripped in SKIP_SYMBOLS:
        return True
    if len(stripped) == 1 and not stripped.isalpha():
        return True
    if EMAIL_RE.fullmatch(stripped):
        return True
    if URL_RE.fullmatch(stripped):
        return True
    if PHONE_RE.fullmatch(stripped):
        return True
    if normalize_acronym(stripped) in KNOWN_ACRONYMS:
        return True
    return bool(NUMBER_ONLY_RE.fullmatch(stripped))


def clean_translation(text):
    """Remove occasional model preambles and return only the translation."""
    if "번역:" in text:
        text = text.rsplit("번역:", 1)[-1].strip()
    for marker in ["번역 결과:", "Translation:", "결과:"]:
        if text.startswith(marker):
            text = text[len(marker) :].strip()
    return text


def build_system_message(matched_terms=None, slide_context=None):
    """Build the system message with optional glossary and slide context."""
    term_hint = ""
    if matched_terms:
        term_hint = "\n\nGlossary terms to preserve:\n" + "\n".join(f"- {term}" for term in matched_terms)

    context_hint = ""
    if slide_context:
        context_hint = (
            "\n\nSlide context for consistency only. Translate only the user text:\n"
            f"{slide_context}"
        )

    return (
        "You are a professional French-to-Korean translator for records-management "
        "and administrative PowerPoint documents.\n"
        "Rules:\n"
        "1. Output Korean only.\n"
        "2. Output only the translation, with no explanations or comments.\n"
        "3. Preserve proper nouns, acronyms, numbers, dates, emails, URLs, and phone numbers.\n"
        "4. Use the provided glossary terms exactly when applicable."
        f"{term_hint}{context_hint}"
    )


def translate_text(text, matched_terms=None, slide_context=None):
    """Translate one text item, using same-slide context when provided."""
    if not text or not text.strip():
        return ""
    if should_skip(text):
        return text

    payload = {
        "model": LLAMA_MODEL,
        "messages": [
            {"role": "system", "content": build_system_message(matched_terms, slide_context)},
            {"role": "user", "content": f"번역할 텍스트:\n\n{text}"},
        ],
        "max_tokens": MAX_TOKENS,
        "temperature": TEMPERATURE,
        "chat_template_kwargs": {"enable_thinking": False},
    }

    try:
        response = requests.post(CHAT_URL, json=payload, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        message = response.json()["choices"][0]["message"]
        return clean_translation((message.get("content") or "").strip())
    except requests.exceptions.ConnectionError:
        print(f"[error] Could not connect to llama.cpp server: {text[:50]}")
    except requests.exceptions.Timeout:
        print(f"[error] Translation timed out: {text[:50]}")
    except Exception as exc:
        print(f"[error] Translation failed for '{text[:50]}': {exc}")
    return ""
