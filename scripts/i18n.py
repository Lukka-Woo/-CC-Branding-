"""
Internationalization module for bilingual presentation support.

Usage in job scripts:
    from scripts.i18n import configure, T, is_bilingual, get_cn, get_en
    configure(lang="bilingual", cn_color="#0E1216", en_color="#8A9199")

Text fields in components accept str | dict{"cn": ..., "en": ...}
"""
from __future__ import annotations
import scripts.brand_tokens as BT

# ── Global config (mutated by configure()) ────────────────────────────────────
LANG         = "cn"                  # "cn" | "en" | "bilingual"
CN_COLOR     = BT.NEUTRAL_900_HEX   # CN body text color (bilingual mode)
EN_COLOR     = BT.NEUTRAL_400_HEX   # EN body text color (bilingual mode)
EN_TITLE_COLOR = BT.NEUTRAL_700_HEX # EN title/subtitle color in bilingual
EN_SZ_RATIO  = 0.85                 # EN font size = CN size × this ratio


def configure(
    lang: str = "cn",
    cn_color: str | None = None,
    en_color: str | None = None,
    en_title_color: str | None = None,
    en_sz_ratio: float | None = None,
):
    """
    Call once at the top of a job script to set the language mode.

    lang        : "cn" | "en" | "bilingual"
    cn_color    : body text color for Chinese content  (bilingual mode)
    en_color    : body text color for English content  (bilingual mode)
    en_title_color : EN subtitle/title color           (bilingual mode)
    en_sz_ratio : EN font size = CN size × ratio       (default 0.85)
    """
    global LANG, CN_COLOR, EN_COLOR, EN_TITLE_COLOR, EN_SZ_RATIO
    LANG = lang
    if cn_color:       CN_COLOR       = cn_color
    if en_color:       EN_COLOR       = en_color
    if en_title_color: EN_TITLE_COLOR = en_title_color
    if en_sz_ratio:    EN_SZ_RATIO    = en_sz_ratio


# ── Text resolvers ────────────────────────────────────────────────────────────

def T(field) -> str:
    """
    Resolve a text field to a single string for current language.
    In bilingual mode, returns the CN string (caller handles EN separately).
    """
    if isinstance(field, str):
        return field
    if isinstance(field, dict):
        if LANG == "en":
            return field.get("en", field.get("cn", ""))
        return field.get("cn", "")
    return str(field) if field is not None else ""


def get_cn(field) -> str:
    if isinstance(field, str):
        return field
    if isinstance(field, dict):
        return field.get("cn", "")
    return ""


def get_en(field) -> str:
    if isinstance(field, dict):
        return field.get("en", "")
    if isinstance(field, str) and LANG == "en":
        return field
    return ""


def is_bilingual() -> bool:
    return LANG == "bilingual"


def has_translation(field) -> bool:
    """True if field carries an EN translation (in bilingual mode)."""
    return is_bilingual() and isinstance(field, dict) and bool(field.get("en"))
