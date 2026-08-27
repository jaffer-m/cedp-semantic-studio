"""Rule-based humanizer — strips AI writing patterns from description text.

No LLM call. Applies vocabulary swaps, filler removal, copula fixes, and
tail-clause removal deterministically. Based on Wikipedia's Signs of AI Writing.
"""

import re

# ── Vocabulary replacement table ──────────────────────────────────────────
# Ordered: longer phrases before shorter words to avoid partial matches.

_VOCAB: list[tuple[str, str]] = [
    # Multi-word phrases first
    ("in order to", "to"),
    ("at this point in time", "now"),
    ("going forward", "from now on"),
    ("it is worth noting that", "note that"),
    ("it is important to note that", "note that"),
    ("it should be noted that", "note that"),
    ("needless to say", ""),
    ("it goes without saying", ""),
    ("as we can see", ""),
    ("in today's rapidly changing world", "today"),
    ("in an era of", "with"),
    ("against this backdrop", ""),
    ("this is a testament to", "this shows"),
    ("this speaks to the importance of", "this highlights"),
    ("not only", ""),
    ("but also", "and"),
    ("rather than", "instead of"),
    ("in connection with", "with"),
    ("associated with", "linked to"),
    ("in association with", "with"),
    ("serves as", "is"),
    ("stands as", "is"),
    ("functions as", "is"),
    ("represents", "is"),
    # Single words
    ("delve into", "explore"),
    ("delves into", "explores"),
    ("delved into", "explored"),
    ("delving into", "exploring"),
    ("tapestry", "mix"),
    ("meticulous", "thorough"),
    ("meticulously", "carefully"),
    ("bolstered", "supported"),
    ("bolster", "support"),
    ("testament", "evidence"),
    ("intricate", "complex"),
    ("intricately", "in a complex way"),
    ("fostering", "building"),
    ("foster", "build"),
    ("showcasing", "showing"),
    ("showcase", "show"),
    ("showcases", "shows"),
    ("showcased", "showed"),
    ("enhance", "improve"),
    ("enhances", "improves"),
    ("enhanced", "improved"),
    ("enhancing", "improving"),
    ("enhancement", "improvement"),
    ("enduring", "lasting"),
    ("underscore", "show"),
    ("underscores", "shows"),
    ("underscored", "showed"),
    ("underscoring", "showing"),
    ("pivotal", "key"),
    ("leverage", "use"),
    ("leverages", "uses"),
    ("leveraged", "used"),
    ("leveraging", "using"),
    ("utilize", "use"),
    ("utilizes", "uses"),
    ("utilized", "used"),
    ("utilizing", "using"),
    ("utilization", "use"),
    ("facilitate", "help"),
    ("facilitates", "helps"),
    ("facilitated", "helped"),
    ("facilitating", "helping"),
    ("facilitate", "enable"),
    ("paramount", "most important"),
    ("robust", "solid"),
    ("comprehensive", "complete"),
    ("navigate", "handle"),
    ("navigates", "handles"),
    ("navigated", "handled"),
    ("navigating", "handling"),
    ("boasts", "has"),
    ("boast", "have"),
    ("nestled", "located"),
    ("vibrant", "active"),
    ("groundbreaking", "new"),
    ("state-of-the-art", "modern"),
    ("cutting-edge", "modern"),
    ("world-class", "high-quality"),
]


def _apply_vocab(text: str) -> str:
    for old, new in _VOCAB:
        pattern = re.compile(r'\b' + re.escape(old) + r'\b', re.IGNORECASE)
        def _replace(m, new=new):
            if not new:
                return ""
            # Preserve leading capitalization
            orig = m.group(0)
            if orig and orig[0].isupper() and new:
                return new[0].upper() + new[1:]
            return new
        text = pattern.sub(_replace, text)
    # Clean up double spaces left by empty replacements
    text = re.sub(r' {2,}', ' ', text).strip()
    return text


# ── Filler tail-clause removal ────────────────────────────────────────────

_TAIL_PATTERNS: list[str] = [
    r',?\s+highlight(?:ing|s) its \w+',
    r',?\s+highlight(?:ing|s) the \w+',
    r',?\s+underscore?(?:ing|s) its \w+',
    r',?\s+underscore?(?:ing|s) the \w+',
    r',?\s+demonstrat(?:ing|es) its \w+',
    r',?\s+demonstrat(?:ing|es) the \w+',
    r',?\s+showcas(?:ing|es) its \w+',
    r',?\s+showcas(?:ing|es) the \w+',
    r',?\s+reflect(?:ing|s) its \w+',
    r',?\s+reflect(?:ing|s) the \w+',
    r',?\s+emphasiz(?:ing|es) its \w+',
    r',?\s+emphasiz(?:ing|es) the \w+',
]

_TAIL_RE = re.compile(
    '(' + '|'.join(_TAIL_PATTERNS) + r')[\w\s]*\.?$',
    re.IGNORECASE,
)


def _remove_tail_clauses(text: str) -> str:
    # Process sentence by sentence
    sentences = re.split(r'(?<=[.!?])\s+', text)
    cleaned = []
    for s in sentences:
        s = _TAIL_RE.sub('', s).strip()
        if s and not s.endswith('.'):
            s += '.'
        if s and s != '.':
            cleaned.append(s)
    return ' '.join(cleaned)


# ── Filler phrase removal ─────────────────────────────────────────────────

_FILLER_PHRASES: list[str] = [
    r'it is worth noting that\s*',
    r'it should be noted that\s*',
    r'it is important to note that\s*',
    r'needless to say,?\s*',
    r'it goes without saying(?: that)?,?\s*',
    r'as we can see,?\s*',
    r'against this backdrop,?\s*',
    r'in today\'s rapidly changing world,?\s*',
    r'in an era of \w+,?\s*',
]

_FILLER_RE = re.compile(
    '|'.join(_FILLER_PHRASES),
    re.IGNORECASE,
)


def _remove_fillers(text: str) -> str:
    text = _FILLER_RE.sub('', text)
    return re.sub(r' {2,}', ' ', text).strip()


# ── Sentence cleanup ──────────────────────────────────────────────────────

def _fix_punctuation(text: str) -> str:
    # Fix ". ." or ".." artifacts
    text = re.sub(r'\.{2,}', '.', text)
    # Fix space before punctuation
    text = re.sub(r'\s+([.,;:])', r'\1', text)
    # Ensure single space after punctuation
    text = re.sub(r'([.!?])([A-Z])', r'\1 \2', text)
    # Capitalize first letter
    if text:
        text = text[0].upper() + text[1:]
    return text.strip()


# ── Public API ────────────────────────────────────────────────────────────

def humanize(text: str) -> str:
    """Apply all humanize rules to a single description string."""
    if not text or not text.strip():
        return text
    text = _remove_fillers(text)
    text = _apply_vocab(text)
    text = _remove_tail_clauses(text)
    text = _fix_punctuation(text)
    return text


def humanize_all(descriptions: dict[str, str]) -> dict[str, str]:
    """Apply humanize() to every value in a {col_name: description} dict."""
    return {k: humanize(v) for k, v in descriptions.items()}
