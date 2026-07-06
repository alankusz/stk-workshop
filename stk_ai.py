"""
Moduł AI — wywołania Claude API do generowania profili kompetencji i pytań STK.
Wzorzec: urllib.request (jak Asystent KFS — omija SIGSEGV w Python 3.9/LibreSSL).
"""
from __future__ import annotations

import json
import re
import time
import urllib.request
import urllib.error
from dataclasses import asdict
from datetime import datetime

from stk_data import (
    CompanyProfile, Competency, CompetencyProfile, STKQuestion, STKOption, STKTest,
)
from stk_prompts import (
    PROFILE_SYSTEM_PROMPT, SINGLE_COMPETENCY_SYSTEM_PROMPT, STK_SYSTEM_PROMPT,
    build_profile_prompt, build_single_competency_prompt, build_stk_prompt,
)

API_URL = "https://api.anthropic.com/v1/messages"
MODEL = "claude-sonnet-4-6"
MAX_TOKENS_PROFILE = 16000
MAX_TOKENS_STK = 16000


# ---------------------------------------------------------------------------
# Wywołanie Claude API
# ---------------------------------------------------------------------------
def call_claude(
    system_prompt: str,
    user_prompt: str,
    api_key: str,
    max_tokens: int = 16000,
    max_retries: int = 4,
) -> str:
    payload = json.dumps({
        "model": MODEL,
        "max_tokens": max_tokens,
        "system": system_prompt,
        "messages": [{"role": "user", "content": user_prompt}],
    }, ensure_ascii=False).encode("utf-8")

    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }

    req = urllib.request.Request(API_URL, data=payload, headers=headers, method="POST")

    for attempt in range(max_retries):
        try:
            with urllib.request.urlopen(req, timeout=240) as response:
                resp = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")
            try:
                err_data = json.loads(body)
                err = err_data.get("error", {})
                err_type = err.get("type", "") if isinstance(err, dict) else ""
                err_msg = err.get("message", body) if isinstance(err, dict) else body
            except json.JSONDecodeError:
                err_type, err_msg = "", body

            if err_type in ("overloaded_error", "rate_limit_error") and attempt < max_retries - 1:
                time.sleep(5 * (attempt + 1))
                continue
            raise RuntimeError(f"API error ({e.code}): {err_msg}")
        except urllib.error.URLError as e:
            raise RuntimeError(f"Blad polaczenia z API: {e.reason}")

        return resp["content"][0]["text"]

    raise RuntimeError("API przeciazone — sprobuj ponownie za chwile.")


# ---------------------------------------------------------------------------
# Parsowanie JSON z odpowiedzi AI
# ---------------------------------------------------------------------------
def _is_closing_quote(s: str, i: int) -> bool:
    """Czy cudzysłów na pozycji i naprawdę zamyka string JSON.

    Zamyka, gdy następny znaczący znak to : } ] lub koniec tekstu.
    Przy przecinku dodatkowo sprawdzamy, czy po nim zaczyna się poprawny
    token JSON — inaczej to cytat w treści (np. mówi "nie", ale...).
    """
    j = i + 1
    while j < len(s) and s[j] in " \t\r\n":
        j += 1
    if j >= len(s) or s[j] in ":}]":
        return True
    if s[j] == ",":
        k = j + 1
        while k < len(s) and s[k] in " \t\r\n":
            k += 1
        return k < len(s) and (s[k] in '"{[-tfn' or s[k].isdigit())
    return False


def _repair_json(s: str) -> str:
    """Escapuje niezaescapowane cudzysłowy wewnątrz stringów JSON.

    Częsty błąd modeli: cytat w treści scenariusza ("...") psuje parsowanie.
    """
    out = []
    in_str = False
    escape = False
    for i, ch in enumerate(s):
        if not in_str:
            if ch == '"':
                in_str = True
            out.append(ch)
            continue
        if escape:
            out.append(ch)
            escape = False
            continue
        if ch == "\\":
            out.append(ch)
            escape = True
            continue
        if ch == '"':
            if _is_closing_quote(s, i):
                in_str = False
                out.append(ch)
            else:
                out.append('\\"')
            continue
        out.append(ch)
    return "".join(out)


def _loads_lenient(s: str) -> dict:
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        return json.loads(_repair_json(s))


def extract_json(text: str) -> dict:
    m = re.search(r"```(?:json)?\s*\n?(.*?)```", text, re.DOTALL)
    if m:
        return _loads_lenient(m.group(1).strip())

    for start_char, end_char in [("{", "}"), ("[", "]")]:
        start = text.find(start_char)
        if start == -1:
            continue
        depth = 0
        in_str = False
        escape = False
        for i in range(start, len(text)):
            ch = text[i]
            if in_str:
                if escape:
                    escape = False
                elif ch == "\\":
                    escape = True
                elif ch == '"':
                    in_str = False
                continue
            if ch == '"':
                in_str = True
            elif ch == start_char:
                depth += 1
            elif ch == end_char:
                depth -= 1
                if depth == 0:
                    return _loads_lenient(text[start : i + 1])
        # nawiasy niedomkniete (np. cudzyslow w tresci) — sprobuj naprawy calosci
        return _loads_lenient(text[start:].strip())

    raise ValueError(f"Nie znaleziono JSON w odpowiedzi AI. Poczatek: {text[:200]}")


def _call_and_parse(
    system_prompt: str,
    user_prompt: str,
    api_key: str,
    max_tokens: int,
    attempts: int = 2,
) -> dict:
    """Wywołuje API i parsuje JSON; przy błędzie parsowania generuje ponownie."""
    last_err: Exception | None = None
    for _ in range(attempts):
        raw = call_claude(system_prompt, user_prompt, api_key, max_tokens)
        try:
            return extract_json(raw)
        except (ValueError, json.JSONDecodeError) as e:
            last_err = e
    raise RuntimeError(
        f"AI zwróciło niepoprawny JSON mimo {attempts} prób. Spróbuj ponownie. ({last_err})"
    )


def _parse_competency(cd: dict) -> Competency:
    """Parsuje dict z odpowiedzi AI do obiektu Competency (skala 1-5)."""
    return Competency(
        name=cd.get("name", ""),
        category=cd.get("category", ""),
        definition=cd.get("definition", ""),
        indicators=cd.get("indicators", []),
        level_1=cd.get("level_1", ""),
        level_2=cd.get("level_2", ""),
        level_3=cd.get("level_3", ""),
        level_4=cd.get("level_4", ""),
        level_5=cd.get("level_5", ""),
    )


# ---------------------------------------------------------------------------
# Generowanie profilu kompetencji (wiele na raz)
# ---------------------------------------------------------------------------
def generate_competency_profile(
    company: CompanyProfile,
    selected_names: list[str],
    api_key: str,
) -> CompetencyProfile:
    """Generuje opisy wybranych kompetencji dla firmy/stanowiska."""
    user_prompt = build_profile_prompt(asdict(company), selected_names)
    data = _call_and_parse(PROFILE_SYSTEM_PROMPT, user_prompt, api_key, MAX_TOKENS_PROFILE)

    profile = CompetencyProfile(company=company)
    items = data.get("competencies", data if isinstance(data, list) else [])
    for cd in items:
        profile.competencies.append(_parse_competency(cd))

    profile.created_at = datetime.now().isoformat()
    return profile


# ---------------------------------------------------------------------------
# Generowanie opisu JEDNEJ kompetencji (dopisywanie własnej)
# ---------------------------------------------------------------------------
def generate_single_competency(
    name: str,
    company: CompanyProfile,
    api_key: str,
) -> Competency:
    """Generuje opis jednej kompetencji w kontekście firmy."""
    company_context = (
        f"Firma: {company.company_name} ({company.industry}, {company.size}). "
        f"Stanowisko: {company.position_name} ({company.position_level}). "
        f"Zadania: {company.key_tasks}."
    )
    user_prompt = build_single_competency_prompt(name, company_context)
    data = _call_and_parse(SINGLE_COMPETENCY_SYSTEM_PROMPT, user_prompt, api_key, 4000)
    return _parse_competency(data)


# ---------------------------------------------------------------------------
# Generowanie pytań STK
# ---------------------------------------------------------------------------
def _fix_question_key(q: STKQuestion) -> bool:
    """Naprawia klucz pytania, gdy da się to zrobić jednoznacznie.

    Scoring best/worst wymaga: 4 opcje, dokładnie jedna z 4 pkt i jedna z 1 pkt.
    Jeśli AI dało np. dwie opcje po 2 pkt i żadnej z 1 pkt, ale najlepsza
    i najgorsza są jednoznaczne (unikalne max i min), przemapowujemy punkty
    na pełną permutację 4-3-2-1 zachowując ranking AI.
    Zwraca True, gdy klucz jest poprawny (po ewentualnej naprawie).
    """
    if len(q.options) != 4:
        return False
    scores = [o.score for o in q.options]
    if sorted(scores) == [1, 2, 3, 4]:
        return True
    hi, lo = max(scores), min(scores)
    if hi == lo or scores.count(hi) != 1 or scores.count(lo) != 1:
        return False  # remis na max lub min — najlepsza/najgorsza niejednoznaczna
    order = sorted(range(4), key=lambda i: scores[i], reverse=True)
    for rank, idx in zip((4, 3, 2, 1), order):
        q.options[idx].score = rank
    return True


def _build_test_from_data(data: dict, company: CompanyProfile, company_context: str) -> STKTest:
    test = STKTest(
        name=f"STK — {company.position_name} ({company.company_name})",
        description=f"Sytuacyjny Test Kompetencyjny dla stanowiska {company.position_name}",
        company_context=company_context,
    )

    seen_ids = set()
    for i, qd in enumerate(data.get("questions", data if isinstance(data, list) else [])):
        qid = str(qd.get("id", "")).strip()
        if not qid or qid in seen_ids:
            qid = f"Q{i + 1}"  # gwarancja unikalnego ID — odpowiedzi sa kluczowane po q.id
        seen_ids.add(qid)
        q = STKQuestion(
            id=qid,
            competency_name=qd.get("competency_name", ""),
            scenario=qd.get("scenario", ""),
        )
        for od in qd.get("options", []):
            q.options.append(STKOption(
                text=od.get("text", ""),
                score=int(od.get("score", 0)),
            ))
        test.questions.append(q)

    test.created_at = datetime.now().isoformat()
    return test


def generate_stk_test(
    competencies: list[Competency],
    company: CompanyProfile,
    api_key: str,
    questions_per_competency: int = 6,
    incidents: list | None = None,
) -> STKTest:
    """Generuje pytania STK.

    incidents: lista dict z CriticalIncident.to_dict() — gdy podane, AI
    opiera dylematy na rzeczywistych zdarzeniach z organizacji.
    """
    comp_dicts = [asdict(c) for c in competencies]
    company_context = (
        f"Firma: {company.company_name} ({company.industry}, {company.size}). "
        f"Stanowisko: {company.position_name} ({company.position_level}). "
        f"Zadania: {company.key_tasks}. "
        f"Kontekst: {company.culture_notes} {company.additional_context}"
    )
    user_prompt = build_stk_prompt(comp_dicts, company_context, questions_per_competency, incidents)

    test = None
    for _ in range(2):
        data = _call_and_parse(STK_SYSTEM_PROMPT, user_prompt, api_key, MAX_TOKENS_STK)
        test = _build_test_from_data(data, company, company_context)
        if test.questions and all(_fix_question_key(q) for q in test.questions):
            return test
        # klucz nie do naprawienia (np. dwie opcje najgorsze) — wygeneruj ponownie

    return test  # ostatnia proba; walidator w aplikacji pokaze ostrzezenie
