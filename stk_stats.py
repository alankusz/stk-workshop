"""
Moduł psychometryczny STK — rzetelność, trafność, moc testu.

Czysty Python (math), bez zależności od scipy — przenośny na Google Apps Script.

Metody i źródła:
- Alfa Cronbacha (Cronbach 1951), progi interpretacji: Nunnally (1978)
- SEM i RCI (Jacobson & Truax 1991) — rzetelna zmiana indywidualna miedzy pomiarami
- Analiza pozycji: trudność + skorygowana korelacja pozycja-wynik (Ebel 1965)
- Korelacja Spearmana przy zmiennych porządkowych (Czakon 2020; konwencja ITET-2)
- Test t dla prób zależnych + d Cohena (dz) — zgodnie z H2 doktoratu ITET-2 v6.0
- Moc testu: aproksymacja normalna (Cohen 1988); N wymagane dla mocy 0.80
- Uwaga metodyczna: alfa bywa zaniżona dla SJT (konstrukty heterogeniczne,
  Catano et al. 2012) — przy niskiej alfie zalecany test-retest.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

from stk_data import STKResult, STKTest


# ---------------------------------------------------------------------------
# Podstawy: statystyki opisowe i rozklad normalny
# ---------------------------------------------------------------------------
def mean(xs: list[float]) -> float:
    return sum(xs) / len(xs) if xs else 0.0


def sample_variance(xs: list[float]) -> float:
    """Wariancja z proby (ddof=1) — standard dla alfy Cronbacha."""
    n = len(xs)
    if n < 2:
        return 0.0
    m = mean(xs)
    return sum((x - m) ** 2 for x in xs) / (n - 1)


def sample_sd(xs: list[float]) -> float:
    return math.sqrt(sample_variance(xs))


def normal_cdf(z: float) -> float:
    return 0.5 * (1.0 + math.erf(z / math.sqrt(2.0)))


def normal_ppf(p: float) -> float:
    """Kwantyl rozkladu normalnego (bisekcja — wystarczajaca dokladnosc)."""
    if p <= 0.0:
        return -8.0
    if p >= 1.0:
        return 8.0
    lo, hi = -8.0, 8.0
    for _ in range(80):
        mid = (lo + hi) / 2
        if normal_cdf(mid) < p:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


# ---------------------------------------------------------------------------
# Rozklad t-Studenta: p-value przez regularyzowana niekompletna funkcje beta
# ---------------------------------------------------------------------------
def _betacf(a: float, b: float, x: float) -> float:
    """Ulamek lancuchowy dla niekompletnej funkcji beta (Numerical Recipes)."""
    MAXIT, EPS, FPMIN = 200, 3.0e-12, 1.0e-300
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c = 1.0
    d = 1.0 - qab * x / qap
    if abs(d) < FPMIN:
        d = FPMIN
    d = 1.0 / d
    h = d
    for m in range(1, MAXIT + 1):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < FPMIN:
            d = FPMIN
        c = 1.0 + aa / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < FPMIN:
            d = FPMIN
        c = 1.0 + aa / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < EPS:
            break
    return h


def betainc_reg(a: float, b: float, x: float) -> float:
    """Regularyzowana niekompletna funkcja beta I_x(a, b)."""
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    ln_bt = (math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b)
             + a * math.log(x) + b * math.log(1.0 - x))
    bt = math.exp(ln_bt)
    if x < (a + 1.0) / (a + b + 2.0):
        return bt * _betacf(a, b, x) / a
    return 1.0 - bt * _betacf(b, a, 1.0 - x) / b


def t_test_p_two_sided(t: float, df: int) -> float:
    """Dwustronne p dla statystyki t o df stopniach swobody."""
    if df <= 0:
        return 1.0
    x = df / (df + t * t)
    return betainc_reg(df / 2.0, 0.5, x)


# ---------------------------------------------------------------------------
# Korelacje: Pearson + Spearman (rangi wiazane -> rangi usrednione)
# ---------------------------------------------------------------------------
def pearson(xs: list[float], ys: list[float]) -> float | None:
    n = len(xs)
    if n < 3 or n != len(ys):
        return None
    mx, my = mean(xs), mean(ys)
    sx = math.sqrt(sum((x - mx) ** 2 for x in xs))
    sy = math.sqrt(sum((y - my) ** 2 for y in ys))
    if sx == 0 or sy == 0:
        return None  # brak zroznicowania
    cov = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    return cov / (sx * sy)


def _ranks(xs: list[float]) -> list[float]:
    """Rangi z usrednianiem przy wiazaniach (ties)."""
    order = sorted(range(len(xs)), key=lambda i: xs[i])
    ranks = [0.0] * len(xs)
    i = 0
    while i < len(order):
        j = i
        while j + 1 < len(order) and xs[order[j + 1]] == xs[order[i]]:
            j += 1
        avg_rank = (i + j) / 2.0 + 1.0
        for k in range(i, j + 1):
            ranks[order[k]] = avg_rank
        i = j + 1
    return ranks


def spearman(xs: list[float], ys: list[float]) -> float | None:
    """Korelacja rang Spearmana — wlasciwa dla zmiennych porzadkowych."""
    if len(xs) < 3 or len(xs) != len(ys):
        return None
    return pearson(_ranks(xs), _ranks(ys))


# ---------------------------------------------------------------------------
# Walidacja poprawnosci testu (klucz odpowiedzi)
# ---------------------------------------------------------------------------
def validate_test(test: STKTest) -> list[str]:
    """Sprawdza strukture testu: kazde pytanie musi miec dokladnie jedna
    opcje najlepsza (4 pkt) i jedna najgorsza (1 pkt) — inaczej scoring
    best/worst jest niejednoznaczny."""
    issues = []
    seen_ids = set()
    for i, q in enumerate(test.questions):
        label = q.id or f"pytanie {i + 1}"
        if not q.id:
            issues.append(f"[{label}] brak ID pytania")
        elif q.id in seen_ids:
            issues.append(f"[{label}] zduplikowane ID pytania")
        seen_ids.add(q.id)

        if not q.scenario.strip():
            issues.append(f"[{label}] pusty scenariusz")
        if len(q.options) != 4:
            issues.append(f"[{label}] liczba opcji: {len(q.options)} (wymagane 4)")

        n_best = sum(1 for o in q.options if o.score == 4)
        n_worst = sum(1 for o in q.options if o.score == 1)
        if n_best != 1:
            issues.append(f"[{label}] opcji z 4 pkt (najlepsza): {n_best} (wymagana dokładnie 1)")
        if n_worst != 1:
            issues.append(f"[{label}] opcji z 1 pkt (najgorsza): {n_worst} (wymagana dokładnie 1)")
    return issues


# ---------------------------------------------------------------------------
# Macierz wynikow: wiersze = uczestnicy, kolumny = pytania (0-2 pkt)
# ---------------------------------------------------------------------------
def build_item_matrix(
    results: list[STKResult], test: STKTest, competency: str | None = None
) -> tuple[list[str], list[list[float]], list[str]]:
    """Zwraca (nazwy uczestnikow, macierz wynikow, ID pytan).

    competency=None -> caly test; inaczej tylko pytania danej kompetencji.
    """
    questions = (test.questions if competency is None
                 else test.questions_for_competency(competency))
    q_ids = [q.id for q in questions]
    names, matrix = [], []
    for r in results:
        by_id = {a.question_id: a.score for a in r.answers}
        if not all(qid in by_id for qid in q_ids):
            continue  # niekompletny wynik — pomijamy
        names.append(r.participant_name)
        matrix.append([float(by_id[qid]) for qid in q_ids])
    return names, matrix, q_ids


# ---------------------------------------------------------------------------
# Rzetelnosc: alfa Cronbacha + SEM
# ---------------------------------------------------------------------------
def cronbach_alpha(matrix: list[list[float]]) -> float | None:
    """Alfa Cronbacha. Wiersze = osoby, kolumny = pozycje. Wymaga N>=2, k>=2."""
    n = len(matrix)
    if n < 2:
        return None
    k = len(matrix[0])
    if k < 2:
        return None
    item_vars = [sample_variance([row[j] for row in matrix]) for j in range(k)]
    totals = [sum(row) for row in matrix]
    total_var = sample_variance(totals)
    if total_var == 0:
        return None
    return (k / (k - 1)) * (1.0 - sum(item_vars) / total_var)


def alpha_interpretation(alpha: float | None) -> str:
    if alpha is None:
        return "brak danych (potrzeba min. 2 uczestników i 2 pozycji o zróżnicowanych wynikach)"
    if alpha >= 0.90:
        return "bardzo wysoka"
    if alpha >= 0.80:
        return "dobra"
    if alpha >= 0.70:
        return "akceptowalna (próg Nunnally 1978)"
    if alpha >= 0.60:
        return "wątpliwa (dopuszczalna w pilotażu)"
    return "niska (uwaga: alfa bywa zaniżona dla SJT — konstrukty heterogeniczne)"


def sem_measurement(sd: float, alpha: float) -> float:
    """Standardowy blad pomiaru: SEM = SD * sqrt(1 - alfa)."""
    return sd * math.sqrt(max(0.0, 1.0 - alpha))


# ---------------------------------------------------------------------------
# Analiza pozycji (trudnosc + moc dyskryminacyjna)
# ---------------------------------------------------------------------------
@dataclass
class ItemStats:
    question_id: str = ""
    competency_name: str = ""
    difficulty: float = 0.0        # sredni wynik / 2 (0-1); wyzsza = latwiejsza
    discrimination: float | None = None  # skorygowana korelacja pozycja-wynik
    flag: str = ""


def item_analysis(
    results: list[STKResult], test: STKTest
) -> list[ItemStats]:
    _, matrix, q_ids = build_item_matrix(results, test)
    comp_by_id = {q.id: q.competency_name for q in test.questions}
    out = []
    if not matrix:
        return out
    k = len(q_ids)
    for j, qid in enumerate(q_ids):
        item_scores = [row[j] for row in matrix]
        rest_totals = [sum(row) - row[j] for row in matrix]  # skorygowany wynik
        diff = mean(item_scores) / 2.0
        disc = pearson(item_scores, rest_totals)
        flag = ""
        if disc is not None and disc < 0.0:
            flag = "UJEMNA dyskryminacja — sprawdz klucz"
        elif disc is not None and disc < 0.20:
            flag = "słaba dyskryminacja (<0.20, Ebel)"
        elif disc is None:
            flag = "brak zróżnicowania odpowiedzi"
        if diff > 0.90:
            flag = (flag + "; " if flag else "") + "bardzo latwa (>0.90)"
        elif diff < 0.10:
            flag = (flag + "; " if flag else "") + "bardzo trudna (<0.10)"
        out.append(ItemStats(
            question_id=qid,
            competency_name=comp_by_id.get(qid, ""),
            difficulty=round(diff, 3),
            discrimination=round(disc, 3) if disc is not None else None,
            flag=flag,
        ))
    return out


# ---------------------------------------------------------------------------
# Trafnosc (aspekt strukturalny): korelacje Spearmana miedzy podskalami
# ---------------------------------------------------------------------------
def subscale_totals(
    results: list[STKResult], test: STKTest
) -> dict[str, list[float]]:
    """Wyniki podskal (kompetencji) per uczestnik — do korelacji miedzy podskalami."""
    out: dict[str, list[float]] = {name: [] for name in test.competency_names}
    for r in results:
        by_comp = r.score_by_competency(test)
        for name in out:
            out[name].append(float(by_comp.get(name, {"score": 0})["score"]))
    return out


def subscale_correlations(
    results: list[STKResult], test: STKTest
) -> list[tuple[str, str, float | None]]:
    """Macierz korelacji Spearmana miedzy podskalami (trafnosc teoretyczna:
    umiarkowane korelacje dodatnie ~0.2-0.6 wspieraja odrebnosc konstruktow;
    ~1.0 = podskale mierza to samo; ~0 lub ujemne = brak spojnosci)."""
    totals = subscale_totals(results, test)
    names = list(totals.keys())
    pairs = []
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            rho = spearman(totals[names[i]], totals[names[j]])
            pairs.append((names[i], names[j], round(rho, 3) if rho is not None else None))
    return pairs


# ---------------------------------------------------------------------------
# Pre/post: test t dla prob zaleznych, d Cohena, RCI, moc testu
# ---------------------------------------------------------------------------
@dataclass
class PairedAnalysis:
    n: int = 0
    mean_pre: float = 0.0
    mean_post: float = 0.0
    sd_pre: float = 0.0
    sd_post: float = 0.0
    mean_diff: float = 0.0
    sd_diff: float = 0.0
    t_stat: float | None = None
    df: int = 0
    p_value: float | None = None
    cohen_dz: float | None = None
    power_posthoc: float | None = None
    rci_per_participant: list[tuple[str, float, str]] = field(default_factory=list)
    sem_pre: float | None = None
    alpha_pre: float | None = None


def cohen_dz_interpretation(d: float) -> str:
    ad = abs(d)
    if ad >= 0.8:
        return "duży"
    if ad >= 0.5:
        return "średni"
    if ad >= 0.2:
        return "mały"
    return "pomijalny"


def power_paired_posthoc(n: int, dz: float, alpha: float = 0.05) -> float:
    """Moc post-hoc testu t dla prob zaleznych (aproksymacja normalna, Cohen 1988)."""
    if n < 2 or dz == 0:
        return alpha
    z_crit = normal_ppf(1.0 - alpha / 2.0)
    nc = abs(dz) * math.sqrt(n)
    return normal_cdf(nc - z_crit) + normal_cdf(-nc - z_crit)


def required_n_paired(dz: float, power: float = 0.80, alpha: float = 0.05) -> int | None:
    """Wymagane N par dla zalozonej mocy (aproksymacja normalna)."""
    if dz == 0:
        return None
    z_a = normal_ppf(1.0 - alpha / 2.0)
    z_b = normal_ppf(power)
    return math.ceil(((z_a + z_b) / abs(dz)) ** 2)


def paired_prepost_analysis(
    pre_results: list[STKResult],
    post_results: list[STKResult],
    test: STKTest,
) -> PairedAnalysis:
    """Analiza sparowana dwoch pomiarow na wynikach lacznych testu.

    Uzycie generyczne: baza -> koncowy (pre->post, post->delayed, pre->delayed).
    Pola *_pre odnosza sie do pomiaru bazowego (pierwszy argument).
    RCI (Jacobson & Truax 1991): (koncowy - baza) / (sqrt(2) * SEM_bazy);
    |RCI| > 1.96 = zmiana rzetelna statystycznie (p < .05).
    """
    pre_by_id = {r.participant_id: r for r in pre_results}
    post_by_id = {r.participant_id: r for r in post_results}
    paired_ids = sorted(set(pre_by_id) & set(post_by_id))

    out = PairedAnalysis()
    pre_totals = [float(pre_by_id[pid].total_score) for pid in paired_ids]
    post_totals = [float(post_by_id[pid].total_score) for pid in paired_ids]
    out.n = len(paired_ids)
    if out.n == 0:
        return out

    out.mean_pre = mean(pre_totals)
    out.mean_post = mean(post_totals)
    out.sd_pre = sample_sd(pre_totals)
    out.sd_post = sample_sd(post_totals)
    diffs = [b - a for a, b in zip(pre_totals, post_totals)]
    out.mean_diff = mean(diffs)
    out.sd_diff = sample_sd(diffs)

    if out.n >= 2 and out.sd_diff > 0:
        out.cohen_dz = out.mean_diff / out.sd_diff
        out.t_stat = out.mean_diff / (out.sd_diff / math.sqrt(out.n))
        out.df = out.n - 1
        out.p_value = t_test_p_two_sided(out.t_stat, out.df)
        out.power_posthoc = power_paired_posthoc(out.n, out.cohen_dz)

    # RCI — wymaga alfy i SD z pomiaru pre (na wszystkich wynikach pre)
    _, pre_matrix, _ = build_item_matrix(pre_results, test)
    out.alpha_pre = cronbach_alpha(pre_matrix)
    if out.alpha_pre is not None and out.sd_pre > 0:
        out.sem_pre = sem_measurement(out.sd_pre, out.alpha_pre)
        s_diff = math.sqrt(2.0) * out.sem_pre
        if s_diff > 0:
            for pid in paired_ids:
                rci = (post_by_id[pid].total_score - pre_by_id[pid].total_score) / s_diff
                verdict = ("poprawa rzetelna" if rci > 1.96
                           else "pogorszenie rzetelne" if rci < -1.96
                           else "zmiana w granicach błędu pomiaru")
                out.rci_per_participant.append(
                    (pre_by_id[pid].participant_name, round(rci, 2), verdict)
                )
    return out


# ---------------------------------------------------------------------------
# Pelny raport psychometryczny (struktura danych dla UI i DOCX)
# ---------------------------------------------------------------------------
@dataclass
class ReliabilityBlock:
    label: str = ""            # "pre" / "post"
    n: int = 0
    k_items: int = 0
    alpha: float | None = None
    alpha_note: str = ""
    sd: float = 0.0
    sem: float | None = None
    subscale_alphas: list[tuple[str, int, float | None]] = field(default_factory=list)


def reliability_block(
    results: list[STKResult], test: STKTest, label: str
) -> ReliabilityBlock:
    names, matrix, q_ids = build_item_matrix(results, test)
    block = ReliabilityBlock(label=label, n=len(names), k_items=len(q_ids))
    if matrix:
        block.alpha = cronbach_alpha(matrix)
        block.alpha_note = alpha_interpretation(block.alpha)
        totals = [sum(row) for row in matrix]
        block.sd = sample_sd(totals)
        if block.alpha is not None:
            block.sem = sem_measurement(block.sd, block.alpha)
        for comp in test.competency_names:
            _, sub_matrix, sub_ids = build_item_matrix(results, test, comp)
            a = cronbach_alpha(sub_matrix) if sub_matrix else None
            block.subscale_alphas.append((comp, len(sub_ids), a))
    return block
