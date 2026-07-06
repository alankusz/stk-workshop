"""
Modele danych dla aplikacji STK (Sytuacyjny Test Kompetencyjny).
JSON-serializable dla przenośności na Google Apps Script.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from typing import Optional


# ---------------------------------------------------------------------------
# Profil firmy i stanowiska
# ---------------------------------------------------------------------------
@dataclass
class CompanyProfile:
    company_name: str = ""
    industry: str = ""
    size: str = ""  # np. "50-250", ">250"
    culture_notes: str = ""  # opis kultury organizacyjnej
    position_name: str = ""
    position_level: str = ""  # np. "specjalista", "kierownik", "dyrektor"
    key_tasks: str = ""  # opis kluczowych zadań
    additional_context: str = ""


# ---------------------------------------------------------------------------
# Skala kompetencji 1-5 (wg Filipowicza, skala rozwojowa)
# ---------------------------------------------------------------------------
LEVEL_LABELS = {
    1: "Niedostateczny",
    2: "Podstawowy",
    3: "Dobry / samodzielny",
    4: "Bardzo dobry",
    5: "Wybitny / ekspert",
}

LEVEL_KEYS = [1, 2, 3, 4, 5]

LEVEL_COLORS = {
    1: "#d32f2f",
    2: "#f57c00",
    3: "#388e3c",
    4: "#1976d2",
    5: "#7b1fa2",
}


# ---------------------------------------------------------------------------
# Katalog kompetencji (Mapa Kompetencji 2025 Filipowicza + uzupełnienia)
# ---------------------------------------------------------------------------
CATEGORY_LABELS = {
    "S": "Społeczne",
    "O": "Osobiste",
    "M": "Menedżerskie",
    "Z": "Specjalistyczne",
}

COMPETENCY_CATALOG: dict[str, list[str]] = {
    "S": [
        "Budowanie relacji",
        "Komunikatywność",
        "Praca zespołowa / współpraca",
        "Rozwiązywanie konfliktów",
        "Współpraca wewnątrz firmy",
        "Wywieranie wpływu",
        "Dzielenie się wiedzą i doświadczeniem",
        "Orientacja na klienta",
        "Identyfikacja z firmą",
        "Negocjowanie",
        "Prezentacja i wystąpienia publiczne",
        "Empatia / inteligencja emocjonalna",
        "Networking",
        "Asertywność",
    ],
    "O": [
        "Podejmowanie decyzji",
        "Rozwiązywanie problemów",
        "Myślenie analityczne",
        "Dążenie do rezultatów / przedsiębiorczość",
        "Samodzielność",
        "Sumienność / rzetelność",
        "Innowacyjność i elastyczność",
        "Rozwój zawodowy / uczenie się",
        "Zarządzanie sobą / efektywność osobista",
        "Odporność na stres / rezyliencja",
        "Zarządzanie czasem",
        "Kreatywność",
        "Adaptacyjność / otwartość na zmiany",
        "Proaktywność / inicjatywa",
    ],
    "M": [
        "Delegowanie",
        "Motywowanie",
        "Przywództwo",
        "Budowanie zespołów",
        "Planowanie",
        "Zarządzanie zespołem",
        "Zarządzanie zmianą",
        "Myślenie strategiczne",
        "Budowanie sprawnej organizacji",
        "Zarządzanie projektami",
        "Zarządzanie procesami",
        "Ocena i rozwój podwładnych",
        "Coaching i mentoring",
        "Zarządzanie talentami",
        "Udzielanie informacji zwrotnej (feedback)",
    ],
    "Z": [
        "Wiedza zawodowa / merytoryczna",
        "Umiejętności techniczne",
        "Orientacja w biznesie",
        "Administracja / prowadzenie dokumentacji",
        "Znajomość języków obcych",
        "Kompetencje cyfrowe / IT",
        "Obsługa klienta (specjalistyczna)",
        "Zarządzanie danymi / analityka",
        "Zarządzanie jakością",
        "Znajomość przepisów i regulacji",
        "Zarządzanie finansami / budżetem",
        "Marketing i sprzedaż",
    ],
}


def get_all_catalog_names() -> list[str]:
    """Zwraca pełną listę nazw kompetencji z katalogu."""
    names = []
    for cat in ["S", "O", "M", "Z"]:
        names.extend(COMPETENCY_CATALOG[cat])
    return names


def get_category_for_name(name: str) -> str:
    """Znajduje kategorię dla nazwy kompetencji z katalogu."""
    for cat, names in COMPETENCY_CATALOG.items():
        if name in names:
            return cat
    return "O"  # domyślnie osobista


# ---------------------------------------------------------------------------
# Kompetencja z 5 poziomami (1-5 wg Filipowicza)
# ---------------------------------------------------------------------------
@dataclass
class Competency:
    name: str = ""
    category: str = ""  # S=społeczna, O=osobista, M=menedżerska, Z=specjalistyczna
    definition: str = ""
    indicators: list[str] = field(default_factory=list)
    level_1: str = ""  # niedostateczny
    level_2: str = ""  # podstawowy
    level_3: str = ""  # dobry / samodzielny
    level_4: str = ""  # bardzo dobry
    level_5: str = ""  # wybitny / ekspert

    def get_level_description(self, level: int) -> str:
        return getattr(self, f"level_{level}", "")


@dataclass
class CompetencyProfile:
    company: CompanyProfile = field(default_factory=CompanyProfile)
    competencies: list[Competency] = field(default_factory=list)
    created_at: str = ""

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2)

    @classmethod
    def from_json(cls, data: str) -> CompetencyProfile:
        d = json.loads(data)
        profile = cls()
        profile.company = CompanyProfile(**d.get("company", {}))
        profile.competencies = [Competency(**c) for c in d.get("competencies", [])]
        profile.created_at = d.get("created_at", "")
        return profile


# ---------------------------------------------------------------------------
# Analiza potrzeb szkoleniowych
# ---------------------------------------------------------------------------
@dataclass
class NeedsAssessmentItem:
    competency_name: str = ""
    current_level: int = 0  # 1-5
    desired_level: int = 0  # 1-5
    importance: int = 3  # 1-5 (ważność kompetencji)
    assessor: str = ""  # "self" / "supervisor" / "trainer"

    @property
    def gap(self) -> int:
        return max(0, self.desired_level - self.current_level)

    @property
    def priority_score(self) -> int:
        return self.gap * self.importance


@dataclass
class NeedsAssessment:
    items: list[NeedsAssessmentItem] = field(default_factory=list)
    assessor_name: str = ""
    assessor_role: str = ""
    notes: str = ""

    def top_priorities(self, n: int = 5) -> list[NeedsAssessmentItem]:
        return sorted(self.items, key=lambda x: x.priority_score, reverse=True)[:n]

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2)

    @classmethod
    def from_json(cls, data: str) -> NeedsAssessment:
        d = json.loads(data)
        assessment = cls(
            assessor_name=d.get("assessor_name", ""),
            assessor_role=d.get("assessor_role", ""),
            notes=d.get("notes", ""),
        )
        assessment.items = [NeedsAssessmentItem(**it) for it in d.get("items", [])]
        return assessment


# ---------------------------------------------------------------------------
# STK — Sytuacyjny Test Kompetencyjny
# ---------------------------------------------------------------------------
@dataclass
class STKOption:
    text: str = ""
    score: int = 0  # 1-4 (1=najgorsza, 4=najlepsza)


@dataclass
class STKQuestion:
    id: str = ""
    competency_name: str = ""
    scenario: str = ""
    options: list[STKOption] = field(default_factory=list)

    @property
    def max_score(self) -> int:
        return max((o.score for o in self.options), default=0)


@dataclass
class STKTest:
    name: str = ""
    description: str = ""
    company_context: str = ""
    questions: list[STKQuestion] = field(default_factory=list)
    created_at: str = ""

    def questions_for_competency(self, name: str) -> list[STKQuestion]:
        return [q for q in self.questions if q.competency_name == name]

    @property
    def max_total_score(self) -> int:
        return len(self.questions) * 2  # max 2 pkt per pytanie (best + worst)

    @property
    def competency_names(self) -> list[str]:
        seen = []
        for q in self.questions:
            if q.competency_name not in seen:
                seen.append(q.competency_name)
        return seen

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2)

    @classmethod
    def from_json(cls, data: str) -> STKTest:
        d = json.loads(data)
        test = cls(
            name=d.get("name", ""),
            description=d.get("description", ""),
            company_context=d.get("company_context", ""),
            created_at=d.get("created_at", ""),
        )
        for qd in d.get("questions", []):
            q = STKQuestion(
                id=qd.get("id", ""),
                competency_name=qd.get("competency_name", ""),
                scenario=qd.get("scenario", ""),
            )
            for od in qd.get("options", []):
                q.options.append(STKOption(**od))
            test.questions.append(q)
        return test


# ---------------------------------------------------------------------------
# Wyniki testu STK
# ---------------------------------------------------------------------------
@dataclass
class STKAnswer:
    """Odpowiedź uczestnika: wskazanie najlepszego i najgorszego zachowania."""
    question_id: str = ""
    best_index: int = -1   # indeks opcji wskazanej jako najlepsza (0-3)
    worst_index: int = -1  # indeks opcji wskazanej jako najgorsza (0-3)
    score: int = 0         # 0, 1 lub 2 (wg Filipowicza)


def score_best_worst(question: STKQuestion, best_idx: int, worst_idx: int) -> int:
    """Scoring wg Filipowicza (s. 41): 0=brak trafień, 1=jedno, 2=oba.

    Klucz: opcja z score=4 to najlepsza, opcja z score=1 to najgorsza.
    Uczestnik wskazuje najlepszą i najgorszą — porównujemy z kluczem.
    """
    if not question.options:
        return 0
    key_best = max(range(len(question.options)), key=lambda i: question.options[i].score)
    key_worst = min(range(len(question.options)), key=lambda i: question.options[i].score)
    hits = 0
    if best_idx == key_best:
        hits += 1
    if worst_idx == key_worst:
        hits += 1
    return hits


@dataclass
class STKResult:
    test_name: str = ""
    participant_name: str = ""
    participant_id: str = ""
    measurement_type: str = ""  # "pre" / "post"
    answers: list[STKAnswer] = field(default_factory=list)
    completed_at: str = ""

    @property
    def total_score(self) -> int:
        return sum(a.score for a in self.answers)

    def score_by_competency(self, test: STKTest) -> dict[str, dict]:
        result = {}
        for name in test.competency_names:
            q_ids = {q.id for q in test.questions_for_competency(name)}
            comp_answers = [a for a in self.answers if a.question_id in q_ids]
            comp_questions = test.questions_for_competency(name)
            score = sum(a.score for a in comp_answers)
            max_score = len(comp_questions) * 2  # max 2 pkt per pytanie
            result[name] = {
                "score": score,
                "max": max_score,
                "pct": round(score / max_score * 100, 1) if max_score > 0 else 0,
            }
        return result

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2)

    @classmethod
    def from_json(cls, data: str) -> STKResult:
        d = json.loads(data)
        result = cls(
            test_name=d.get("test_name", ""),
            participant_name=d.get("participant_name", ""),
            participant_id=d.get("participant_id", ""),
            measurement_type=d.get("measurement_type", ""),
            completed_at=d.get("completed_at", ""),
        )
        for ad in d.get("answers", []):
            result.answers.append(STKAnswer(**ad))
        return result


# Porownania miedzy pomiarami (pre/post/delayed) licza:
# - dashboard (stk_app.py, zakladka 4) — na bazie score_by_competency()
# - psychometria (stk_stats.paired_prepost_analysis)
