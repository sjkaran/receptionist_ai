"""Text normalisation, light stemming, synonym folding and small-talk intents.

Everything here is dependency-free and deterministic, so retrieval stays fast
and explainable.
"""
from __future__ import annotations

import re
import unicodedata
from collections import Counter

STOPWORDS = frozenset(
    """a an the is are was were be been am do does did can could would will shall should may might
    i me my we our us you your yours it its this that these those of to in on at for with about from
    by as please tell know want need like hi hello hey there some any so just also and or but if then
    than what whats how let get show kindly sir madam have has available""".split()
)

# Words in one group are treated as the same concept (first word = canonical form).
SYNONYM_GROUPS = [
    "hour hours timing timings time open opening close closing",
    "location address where located directions reach find map",
    "price cost fee fees charge charges rate rates pricing much",
    "phone call contact number telephone mobile",
    "email mail",
    "appointment book booking meeting meet reserve reservation",
    "restroom toilet washroom bathroom lavatory",
    "wifi internet wireless password",
    "parking park",
    "manager boss director head owner",
    "job career vacancy vacancies hiring recruitment",
    "service services offer provide product products give",
    "refund return cancel cancellation",
    "delivery shipping courier",
    "visit visitor guest",
]
_CANON: dict[str, str] = {}
for _g in SYNONYM_GROUPS:
    _words = _g.split()
    for _w in _words:
        _CANON.setdefault(_w, _words[0])

_NON_ALNUM = re.compile(r"[^a-z0-9]+")


def normalize(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "").encode("ascii", "ignore").decode("ascii")
    return _NON_ALNUM.sub(" ", text.lower().replace("'", "")).strip()


def _stem(w: str) -> str:
    if len(w) <= 3:
        return w
    for suf, rep in (("ies", "y"), ("sses", "ss"), ("ing", ""), ("ed", ""), ("es", ""), ("s", ""), ("ly", "")):
        if w.endswith(suf) and len(w) - len(suf) >= 3:
            return w[: -len(suf)] + rep
    return w


def term_list(text: str) -> list[str]:
    out = []
    for w in normalize(text).split():
        if w in STOPWORDS:
            continue
        s = _stem(w)
        out.append(_CANON.get(w) or _CANON.get(s) or s)
    return out


def terms(text: str) -> Counter:
    return Counter(term_list(text))


def trigrams(text: str) -> set[str]:
    s = f" {normalize(text)} "
    return {s[i : i + 3] for i in range(len(s) - 2)}


def dice(a: set, b: set) -> float:
    if not a or not b:
        return 0.0
    return 2.0 * len(a & b) / (len(a) + len(b))


# --------------------------------------------------------------------------
# Small-talk / dialog intents (operate on normalised text)
# --------------------------------------------------------------------------
_SMALLTALK = [
    ("greeting", re.compile(r"^(hi|hello|hey|hiya|namaste|good morning|good afternoon|good evening)( \w+)?$")),
    ("thanks", re.compile(r"^(thanks|thank you|thankyou|thx|cheers|much appreciated|great thanks)\b")),
    ("bye", re.compile(r"^(bye|goodbye|good bye|see you|see ya|take care|thats all|that is all|nothing else)\b")),
    ("how_are_you", re.compile(r"\bhow are you\b|\bhows it going\b")),
    ("identity", re.compile(r"\b(who|what) are you\b|\byour name\b|\bwho am i (talking|speaking) (to|with)\b")),
]
_MESSAGE_INTENT = re.compile(
    r"\b(leave|take|record|send|give|pass)\b.*\b(message|note)\b|\bmessage (for|to)\b|\bleave a message\b"
)
_CANCEL = re.compile(r"\b(cancel|never ?mind|forget it|stop|no thanks)\b")
_YES = re.compile(r"^(yes|yeah|yep|yup|sure|ok|okay|please|go ahead|definitely|of course|why not|alright)\b")
_NO = re.compile(r"^(no|nope|nah|not now|no thanks|dont|do not)\b")


def smalltalk_intent(text: str) -> str | None:
    n = normalize(text)
    if not n or len(n.split()) > 6:
        return None
    for name, rx in _SMALLTALK:
        if rx.search(n):
            return name
    return None


def wants_message(text: str) -> bool:
    return bool(_MESSAGE_INTENT.search(normalize(text)))


def is_cancel(text: str) -> bool:
    return bool(_CANCEL.search(normalize(text)))


def is_yes(text: str) -> bool:
    return bool(_YES.search(normalize(text)))


def is_no(text: str) -> bool:
    return bool(_NO.search(normalize(text)))


_NAME_PREFIX = re.compile(r"^(my name is|my name's|i am|i'm|im|it is|it's|its|this is|call me|name is)\s+", re.I)


def parse_name(text: str) -> str:
    t = _NAME_PREFIX.sub("", text.strip())
    t = re.sub(r"[^\w\s.'-]", "", t).strip()
    return " ".join(w.capitalize() for w in t.split()[:4])[:40]


def clean_for_speech(text: str, max_chars: int = 420) -> str:
    """Strip markdown / list artefacts an LLM may emit and keep replies short."""
    t = re.sub(r"[*_`#>]+", "", text or "")
    t = re.sub(r"^\s*(?:[-•]|\d+[.)])\s*", "", t, flags=re.M)
    t = re.sub(r"\s+", " ", t).strip().strip('"')
    if len(t) > max_chars:
        cut = t[:max_chars]
        end = max(cut.rfind(". "), cut.rfind("! "), cut.rfind("? "))
        t = cut[: end + 1] if end > 80 else cut.rstrip() + "…"
    return t
