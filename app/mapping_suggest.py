from __future__ import annotations

import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Iterable


_NON_ALNUM = re.compile(r"[^a-z0-9]+")


def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    s = _NON_ALNUM.sub(" ", s)
    s = " ".join(s.split())
    return s


def _token_set(s: str) -> set[str]:
    return set(_norm(s).split())


def _seq_ratio(a: str, b: str) -> float:
    a_n = _norm(a)
    b_n = _norm(b)
    if not a_n or not b_n:
        return 0.0
    return SequenceMatcher(a=a_n, b=b_n).ratio()


def _jaccard(a: str, b: str) -> float:
    ta = _token_set(a)
    tb = _token_set(b)
    if not ta or not tb:
        return 0.0
    inter = len(ta & tb)
    union = len(ta | tb)
    return inter / union if union else 0.0


def _synonym_boost(target: str, source: str) -> float:
    t = _norm(target)
    s = _norm(source)
    pairs = [
        ({"phone", "mobile", "contact", "cell"}, {"phone", "mobile", "contact", "cell", "number", "no"}),
        ({"dob", "birth", "birthday", "date"}, {"dob", "birth", "birthday", "date"}),
        ({"pin", "pincode", "zip", "postal"}, {"pin", "pincode", "zip", "postal", "code"}),
        ({"email", "mail"}, {"email", "mail"}),
        ({"state", "province"}, {"state", "province"}),
        ({"city", "town"}, {"city", "town"}),
        ({"name", "full"}, {"name", "full"}),
    ]
    for tset, sset in pairs:
        if any(w in t.split() for w in tset) and any(w in s.split() for w in sset):
            return 0.08
    return 0.0


@dataclass(frozen=True)
class Suggestion:
    source_header: str
    score: float


def suggest_sources_for_target(
    target_header: str, source_headers: Iterable[str], top_k: int = 3
) -> list[Suggestion]:
    scored: list[Suggestion] = []
    for sh in source_headers:
        score = 0.62 * _seq_ratio(target_header, sh) + 0.38 * _jaccard(target_header, sh)
        score += _synonym_boost(target_header, sh)
        scored.append(Suggestion(source_header=sh, score=float(max(0.0, min(1.0, score)))))
    scored.sort(key=lambda x: x.score, reverse=True)
    return scored[: max(1, top_k)]

