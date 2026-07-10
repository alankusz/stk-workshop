"""
Prompty AI dla generowania profili kompetencji i pytań STK.
Few-shot examples oparte na Mapie Kompetencji 2025 (Filipowicz).
Skala: 1-5 (1=niedostateczny, 5=wybitny).
"""

# ---------------------------------------------------------------------------
# Przykłady kompetencji z Mapy Kompetencji 2025 (do few-shot)
# ---------------------------------------------------------------------------
COMPETENCY_EXAMPLES = """
### Przykład 1: Kompetencja społeczna

Nazwa: Budowanie relacji
Kategoria: S (społeczna)
Definicja: Nawiązywanie i budowanie dobrych relacji. Dbałość o kontakty i efektywną współpracę. Pozyskiwanie zaufania i poparcia współpracowników.
Wskaźniki behawioralne:
- Nawiązywanie relacji
- Rozwijanie/budowanie relacji - dbałość o kontakty
- Skuteczne rozwiązywanie nieporozumień
- Pozyskiwanie zaufania i poparcia
- Udzielanie wsparcia

Poziom 1 (niedostateczny): Jest zamknięty w sobie. Nie utrzymuje regularnych kontaktów z innymi, nie oddzwania, unika spotkań. W razie wystąpienia nieporozumień unika działań lub reaguje nieadekwatnie. Dba o swoje interesy za wszelką cenę. Odmawia pomocy innym.

Poziom 2 (podstawowy): W razie potrzeby nawiązuje kontakt, ale rzadko sam wykazuje inicjatywę. Utrzymuje kontakt, lecz czasem zapomina o ważnych telefonach. Podejmuje działania przy nieporozumieniach, ale nie zawsze skutecznie. Stara się pozyskać zaufanie, ale priorytetem są własne cele. Pomaga na prośbę, nie z inicjatywy.

Poziom 3 (dobry/samodzielny): Łatwo i z własnej inicjatywy nawiązuje kontakty. Utrzymuje stały kontakt bezpośredni i pośredni. Właściwie analizuje nieporozumienia i skutecznie je rozwiązuje. Buduje zaufanie na bazie wiarygodności. Chętnie pomaga i przyjmuje pomoc.

Poziom 4 (bardzo dobry): Dobrze nawiązuje kontakt nawet z osobami niechętnymi. Aktywnie dba o trwały kontakt — inicjuje spotkania, regularnie dzwoni. Spokojnie radzi sobie z nieporozumieniami, wykorzystując je do wzmacniania relacji. Jest godny zaufania nawet przy sprzecznych interesach. Zawsze oferuje wsparcie.

Poziom 5 (wybitny): Skutecznie przełamuje lody bez względu na nastawienie i okoliczności. Inicjuje przedsięwzięcia pogłębiające relacje. Przewiduje nieporozumienia i aktywnie im przeciwdziała. Wdraża wartości wspierające budowanie relacji. Mobilizuje i organizuje zasoby, by pomagać innym; propaguje tę postawę.

### Przykład 2: Kompetencja menedżerska

Nazwa: Delegowanie
Kategoria: M (menedżerska)
Definicja: Przekazywanie pracownikom zadań do wykonania. Zapewnienie odpowiednich informacji i uprawnień oraz wsparcia niezbędnego do ich realizacji.
Wskaźniki behawioralne:
- Określenie zadań do wykonania przez pracowników
- Dobór pracowników do zadań
- Zapewnienie odpowiednich informacji
- Przekazywanie uprawnień
- Wsparcie w realizacji trudnych zadań

Poziom 1 (niedostateczny): Ogólnie formułuje cel, czasami nie określa rezultatów. Nie zna kompetencji pracowników — przydziela za trudne lub za łatwe zadania. Wyjaśnia niezrozumiale. Unika przekazywania uprawnień. Deleguje tylko najprostsze zadania.

Poziom 2 (podstawowy): Jasno formułuje cel, nie zawsze precyzyjnie określa rezultaty. Zna kompetencje, ale nie zawsze trafnie dobiera osoby. Używa ogólnikowych określeń. Przekazuje uprawnienia, ale czasem nieadekwatne. Deleguje trudne zadania bez odpowiedniego wsparcia.

Poziom 3 (dobry/samodzielny): Precyzyjnie definiuje cel i oczekiwane rezultaty. Trafnie dopasowuje zadania do kompetencji pracowników. Przekazuje zadania jasno z konkretnymi wskazówkami. Adekwatnie przekazuje uprawnienia. Zapewnia wsparcie w trudnych sytuacjach.

Poziom 4 (bardzo dobry): Definiuje cel w kontekście celów zespołu/firmy. Wykorzystuje delegowanie do rozwoju pracowników. Wyjaśnia nie tylko co, ale i dlaczego. Zachęca do samodzielnego podejmowania decyzji. Wspiera nawet w skomplikowanych zadaniach, nie przejmując ich.

Poziom 5 (wybitny): Tworzy systemy ułatwiające precyzyjne delegowanie w organizacji. Buduje kulturę odpowiedzialności i autonomii. Wdraża rozwiązania zapewniające przepływ informacji przy delegowaniu. Przekazuje uprawnienia strategiczne. Mentoruje innych liderów w zakresie delegowania.

### Przykład 3: Kompetencja osobista

Nazwa: Podejmowanie decyzji
Kategoria: O (osobista)
Definicja: Podejmowanie trafnych i terminowych decyzji. Branie odpowiedzialności za podjęte decyzje i ich konsekwencje.
Wskaźniki behawioralne:
- Analiza sytuacji przed podjęciem decyzji
- Terminowość podejmowania decyzji
- Trafność decyzji
- Branie odpowiedzialności za decyzje
- Komunikowanie decyzji

Poziom 1 (niedostateczny): Unika podejmowania decyzji lub odwleka je. Podejmuje decyzje bez analizy sytuacji. Nie bierze odpowiedzialności za konsekwencje. Nie komunikuje decyzji zainteresowanym.

Poziom 2 (podstawowy): Podejmuje decyzje w standardowych sytuacjach, potrzebuje wsparcia w trudniejszych. Analizuje dostępne informacje, ale nie zawsze dostrzega wszystkie czynniki. Konsultuje trudne decyzje z przełożonym. Komunikuje decyzje, ale nie zawsze z uzasadnieniem.

Poziom 3 (dobry/samodzielny): Samodzielnie podejmuje decyzje w swoim obszarze, opierając się na faktach i procedurach. Racjonalnie argumentuje swoje decyzje. Podejmuje decyzje w terminie, uwzględniając interesy stron. Prezentuje decyzje jasno i zrozumiale.

Poziom 4 (bardzo dobry): Samodzielnie podejmuje decyzje nawet bez wyczerpujących informacji. Bierze odpowiedzialność za dobre i błędne decyzje. Podejmuje trudne decyzje pod presją czasu. Uzasadnia kontrowersyjne decyzje zainteresowanym.

Poziom 5 (wybitny): Trafnie podejmuje decyzje obarczone dużym ryzykiem. Zachęca innych do samodzielnego podejmowania decyzji. Inicjuje procesy usprawniające podejmowanie decyzji w organizacji. Stanowi autorytet i wzór dla innych.
"""

# ---------------------------------------------------------------------------
# System prompt: Generowanie profilu kompetencji
# ---------------------------------------------------------------------------
PROFILE_SYSTEM_PROMPT = """Jesteś ekspertem w zarządzaniu kompetencjami zawodowymi z 20-letnim doświadczeniem we wdrażaniu modeli kompetencyjnych w polskich firmach. Tworzysz opisy kompetencji zgodnie z metodyką Grzegorza Filipowicza (skala rozwojowa 5-poziomowa 1-5).

ZASADY TWORZENIA OPISU KOMPETENCJI:

1. DEFINICJA: Zwięzła (1-2 zdania), opisuje istotę kompetencji w kontekście stanowiska.

2. WSKAŹNIKI BEHAWIORALNE: 4-6 konkretnych, obserwowalnych zachowań. NIE używaj: "ma świadomość", "rozumie" — używaj: "potrafi", "realizuje", "wdraża".

3. PIĘĆ POZIOMÓW (1-5):
   - 1 (niedostateczny): Brak kompetencji lub zachowania szkodliwe. Osoba nie realizuje nawet podstawowych zadań w tym obszarze.
   - 2 (podstawowy/uczący się): Realizuje zadania w prostych sytuacjach, potrzebuje wsparcia w złożonych. Popełnia błędy.
   - 3 (dobry/samodzielny): Samodzielnie realizuje zadania na wymaganym poziomie. To jest NORMA dla danego stanowiska.
   - 4 (bardzo dobry): Wspiera innych, radzi sobie w trudnych sytuacjach, przekracza standardy. Wzór w zespole.
   - 5 (wybitny/ekspert): Tworzy systemy, wdraża innowacje, mentoruje. Wpływa na całą organizację.

4. KLUCZOWE REGUŁY:
   - Każdy poziom musi opisywać KONKRETNE ZACHOWANIA, nie ogólniki
   - Poziomy muszą się jakościowo różnić (nie tylko "lepiej/gorzej")
   - Opisy muszą być dopasowane do kontekstu firmy i stanowiska
   - Kategoryzuj: S=społeczna, O=osobista, M=menedżerska, Z=specjalistyczna
   - Opisy na każdym poziomie: 2-4 zdania

5. FORMAT ODPOWIEDZI: Zwróć TYLKO czysty JSON (bez markdown, bez ```). Klucz "competencies" z listą obiektów: name, category, definition, indicators (lista stringów), level_1, level_2, level_3, level_4, level_5.

6. WAŻNE — POPRAWNOŚĆ JSON: Wewnątrz tekstów NIE używaj prostego cudzysłowu ("). Cytaty zapisuj polskimi cudzysłowami „..." albo apostrofami '...'.
"""

PROFILE_EXAMPLES_NOTE = """
Poniżej przykłady prawidłowo opisanych kompetencji (z Mapy Kompetencji 2025, Filipowicz) jako WZÓR jakości i formatu:
""" + COMPETENCY_EXAMPLES

# ---------------------------------------------------------------------------
# System prompt: Generowanie opisu JEDNEJ kompetencji (dla "dopisz własną")
# ---------------------------------------------------------------------------
SINGLE_COMPETENCY_SYSTEM_PROMPT = """Jesteś ekspertem w zarządzaniu kompetencjami zawodowymi. Stwórz pełny opis jednej kompetencji zgodnie z metodyką Filipowicza (skala 1-5).

FORMAT ODPOWIEDZI: Zwróć TYLKO czysty JSON (bez markdown, bez ```) z polami: name, category, definition, indicators (lista), level_1, level_2, level_3, level_4, level_5.

Każdy poziom (1-5): 2-4 zdania opisujące konkretne zachowania.
- 1 = niedostateczny (brak kompetencji, zachowania szkodliwe)
- 2 = podstawowy (potrzebuje wsparcia, popełnia błędy)
- 3 = dobry/samodzielny (NORMA — realizuje samodzielnie)
- 4 = bardzo dobry (wspiera innych, przekracza standardy)
- 5 = wybitny (tworzy systemy, mentoruje, wpływa na organizację)
"""

# ---------------------------------------------------------------------------
# System prompt: Generowanie pytań STK
# ---------------------------------------------------------------------------
STK_SYSTEM_PROMPT = """Jesteś ekspertem w tworzeniu Sytuacyjnych Testów Kompetencyjnych (STK / SJT — Situational Judgment Tests). Tworzysz dylematy sytuacyjne do pomiaru kompetencji pracowników PRZED i PO szkoleniu.

ZASADY TWORZENIA PYTAŃ STK:

1. DYLEMAT SYTUACYJNY (scenariusz):
   - 3-5 zdań opisujących realistyczną sytuację z pracy
   - Dopasowany do kontekstu firmy i stanowiska
   - Zawiera konflikt, presję lub konieczność podjęcia decyzji
   - Nie sugeruje "poprawnej" odpowiedzi
   - Optymalnie 600-800 znaków

2. CZTERY OPCJE ODPOWIEDZI:
   - Każda opisuje konkretne ZACHOWANIE (nie opinię)
   - Max 2 zdania na opcję (~120 znaków)
   - Opcje muszą być realistyczne — żadna nie powinna być absurdalna
   - Nie używaj "mimo ryzyka", "starasz się" — te sformułowania sugerują odpowiedź
   - Opcje NIE są ułożone od najgorszej do najlepszej

3. SCORING (klucz odpowiedzi):
   - 4 = zachowanie najbardziej kompetentne (odpowiada poziomowi 4-5)
   - 3 = zachowanie dobre (odpowiada poziomowi 3)
   - 2 = zachowanie słabe (odpowiada poziomowi 2)
   - 1 = zachowanie najsłabsze (odpowiada poziomowi 1)
   - KRYTYCZNE: w każdym pytaniu użyj KAŻDEJ wartości dokładnie raz — jedna opcja 4, jedna 3, jedna 2, jedna 1 (pełna permutacja). Scoring best/worst wymaga dokładnie jednej opcji najlepszej (4) i dokładnie jednej najgorszej (1).
   - Kolejność opcji MUSI być losowa (nie od 4 do 1)

4. FORMAT INSTRUKCJI: Behawioralny ("Co byś zrobił w tej sytuacji?")

5. ZRÓŻNICOWANIE:
   - Pytania dla jednej kompetencji muszą dotyczyć RÓŻNYCH sytuacji
   - Progresja trudności: od standardowych sytuacji po złożone

6. FORMAT ODPOWIEDZI: Zwróć TYLKO czysty JSON (bez markdown, bez ```) z kluczem "questions" — lista obiektów: id, competency_name, scenario, options (lista: text, score).

7. WAŻNE — POPRAWNOŚĆ JSON: Wewnątrz tekstów (scenario, text) NIE używaj prostego cudzysłowu ("). Cytaty i wypowiedzi zapisuj polskimi cudzysłowami „..." albo apostrofami '...'.
"""


# ---------------------------------------------------------------------------
# Funkcje tworzące pełne prompty
# ---------------------------------------------------------------------------
def _format_tasks_for_prompt(tasks_list: list) -> str:
    """Formatuje liste zadan jako czytelna tabele do promptu AI."""
    if not tasks_list:
        return "nie podano"
    lines = ["Analiza pracy (skala 1-3):"]
    for t in tasks_list:
        name = t.get("name", "").strip()
        if not name:
            continue
        f = t.get("frequency", 2)
        i = t.get("importance", 2)
        d = t.get("difficulty", 2)
        prio = "Wysoki" if i == 3 else ("Sredni" if i == 2 else "Niski")
        lines.append(f"  - {name} | czestotliwosc:{f}/3 | waznosc:{i}/3 | trudnosc:{d}/3 | priorytet:{prio}")
    lines.append("Zadania o wysokim priorytecie wyznaczaja kompetencje kluczowe.")
    return "\n".join(lines)


def build_profile_prompt(company: dict, selected_names: list[str]) -> str:
    """Buduje user prompt do generowania profilu kompetencji."""
    names_str = "\n".join(f"- {n}" for n in selected_names) if selected_names else "Brak — dobierz sam."
    tasks_list = company.get("key_tasks_list", [])
    if tasks_list:
        tasks_text = _format_tasks_for_prompt(tasks_list)
    else:
        tasks_text = company.get("key_tasks", "nie podano") or "nie podano"
    return f"""Stwórz opisy kompetencji na 5 poziomach (1-5) dla poniższych kompetencji w kontekście firmy i stanowiska.

FIRMA:
- Nazwa: {company.get('company_name', 'nie podano')}
- Branża: {company.get('industry', 'nie podano')}
- Wielkość: {company.get('size', 'nie podano')}
- Kultura/kontekst: {company.get('culture_notes', 'nie podano')}

STANOWISKO:
- Nazwa: {company.get('position_name', 'nie podano')}
- Poziom: {company.get('position_level', 'nie podano')}
- Kluczowe zadania:
{tasks_text}
- Dodatkowy kontekst: {company.get('additional_context', 'nie podano')}

KOMPETENCJE DO OPISANIA:
{names_str}

{PROFILE_EXAMPLES_NOTE}

Wygeneruj opisy w formacie JSON. Pamiętaj:
- Opisy poziomów muszą być dopasowane do kontekstu firmy i branży
- Dla każdej kompetencji: definition, indicators (4-6), level_1 do level_5
- Przypisz kategorię: S/O/M/Z"""


def build_single_competency_prompt(name: str, company_context: str) -> str:
    """Buduje prompt do generowania opisu JEDNEJ kompetencji."""
    return f"""Stwórz pełny opis kompetencji "{name}" na 5 poziomach (1-5).

Kontekst firmy/stanowiska: {company_context}

{PROFILE_EXAMPLES_NOTE}

Zwróć JSON z polami: name, category (S/O/M/Z), definition, indicators, level_1, level_2, level_3, level_4, level_5."""


def build_stk_prompt(
    competencies: list,
    company_context: str,
    questions_per_competency: int = 6,
    incidents=None,
) -> str:
    """Buduje user prompt do generowania pytań STK.

    incidents: lista słowników CriticalIncident.to_dict() — opcjonalne;
    gdy podane, AI opiera dylematy na rzeczywistych zdarzeniach z organizacji.
    """
    comp_list = ""
    for c in competencies:
        comp_list += f"\n### {c['name']} ({c.get('category', '?')})\n"
        comp_list += f"Definicja: {c.get('definition', '')}\n"
        indicators = c.get("indicators", [])
        if indicators:
            comp_list += "Wskaźniki: " + "; ".join(indicators) + "\n"
        comp_list += f"Poziom 3 (norma): {c.get('level_3', '')}\n"
        comp_list += f"Poziom 4 (bardzo dobry): {c.get('level_4', '')}\n"

    incidents_section = ""
    if incidents:
        incidents_section = "\n\nINCYDENTY KRYTYCZNE (rzeczywiste sytuacje z organizacji — opieraj na nich dylematy):\n"
        # Grupuj incydenty per kompetencja
        by_comp: dict[str, list[dict]] = {}
        for inc in incidents:
            cname = inc.get("competency_name", "Inne")
            by_comp.setdefault(cname, []).append(inc)
        for cname, incs in by_comp.items():
            incidents_section += f"\n#### Kompetencja: {cname}\n"
            for j, inc in enumerate(incs, 1):
                incidents_section += f"Incydent {j}:\n"
                if inc.get("situation"):
                    incidents_section += f"  Sytuacja: {inc['situation']}\n"
                if inc.get("actors"):
                    incidents_section += f"  Zaangażowani: {inc['actors']}\n"
                if inc.get("action"):
                    incidents_section += f"  Decyzja/działanie: {inc['action']}\n"
                if inc.get("reasoning"):
                    incidents_section += f"  Powód: {inc['reasoning']}\n"
                if inc.get("result"):
                    incidents_section += f"  Rezultat: {inc['result']}\n"
                if inc.get("best_alternative"):
                    incidents_section += f"  Najlepsza reakcja: {inc['best_alternative']}\n"
                if inc.get("worst_alternative"):
                    incidents_section += f"  Najgorsza reakcja: {inc['worst_alternative']}\n"
        incidents_section += (
            "\nINSTRUKCJA: Przekształć powyższe incydenty w dylematy sytuacyjne — "
            "odanominizuj szczegóły (zmień imiona, daty, szczegóły identyfikujące), "
            "ale zachowaj istotę sytuacji i charakter decyzji. "
            "Jeśli incydent zawiera już najlepszą/najgorszą reakcję, użyj ich jako opcji scoringowych."
        )

    return f"""Wygeneruj Sytuacyjny Test Kompetencyjny (STK) — {questions_per_competency} pytań na każdą kompetencję.

KONTEKST FIRMY/STANOWISKA:
{company_context}

KOMPETENCJE DO POMIARU:
{comp_list}{incidents_section}

WYMAGANIA:
- {questions_per_competency} dylematów sytuacyjnych na KAŻDĄ kompetencję
- Każdy dylemat: sytuacja (3-5 zdań) + 4 opcje odpowiedzi z scoringiem 1-4 (każda wartość dokładnie raz: 4, 3, 2, 1)
- Opcje w LOSOWEJ kolejności (nie od najlepszej do najgorszej)
- Dylematy dopasowane do kontekstu firmy i stanowiska
- ID pytań: K1_1, K1_2... (K=numer kompetencji, _numer pytania)
- Format: czysty JSON z kluczem "questions"

Zwróć wyłącznie JSON."""
