"""
Asystent STK — Aplikacja Warsztatowa
Budowanie profilu kompetencji i analiza potrzeb szkoleniowych.
Wynik: JSON (import do Asystenta STK Trenera) + XLSX (czytelny raport).
"""
from __future__ import annotations

import io
import json
import os

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from stk_data import (
    CompanyProfile, CompetencyProfile, NeedsAssessmentItem, NeedsAssessment,
    LEVEL_LABELS, LEVEL_KEYS, LEVEL_COLORS,
    COMPETENCY_CATALOG, CATEGORY_LABELS,
)
from stk_ai import generate_competency_profile, generate_single_competency


# ---------------------------------------------------------------------------
# Eksport XLSX
# ---------------------------------------------------------------------------
_NAVY = "1B4F8A"
_LIGHT = "E8F0FB"
_RED   = "FFCCCC"
_AMBER = "FFF0CC"
_GREEN = "CCFFCC"


def _hdr(ws, row: int, col: int, text: str) -> None:
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    c.fill = PatternFill("solid", fgColor=_NAVY)
    c.alignment = Alignment(wrap_text=True, vertical="top")


def _cell(ws, row: int, col: int, value, bold: bool = False, bg: str | None = None) -> None:
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Arial", bold=bold, size=10)
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(wrap_text=True, vertical="top")


def build_workshop_xlsx(
    profile: CompetencyProfile,
    assessment: NeedsAssessment | None,
) -> bytes:
    wb = Workbook()

    # ------------------------------------------------------------------ #
    # Arkusz 1: Profil kompetencji
    # ------------------------------------------------------------------ #
    ws1 = wb.active
    ws1.title = "Profil_kompetencji"

    ws1.merge_cells("A1:I1")
    t = ws1.cell(row=1, column=1,
                 value=f"Profil kompetencji — {profile.company.position_name} ({profile.company.company_name})")
    t.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=_NAVY)
    t.alignment = Alignment(horizontal="center")
    ws1.row_dimensions[1].height = 24

    info = [
        ("Firma",             profile.company.company_name),
        ("Branża",            profile.company.industry),
        ("Wielkość",          profile.company.size),
        ("Stanowisko",        profile.company.position_name),
        ("Poziom",            profile.company.position_level),
        ("Kluczowe zadania",  profile.company.key_tasks),
        ("Kultura / kontekst", profile.company.culture_notes),
    ]
    for i, (label, val) in enumerate(info, start=2):
        _cell(ws1, i, 1, label, bold=True, bg=_LIGHT)
        ws1.merge_cells(f"B{i}:I{i}")
        _cell(ws1, i, 2, val)

    hr = len(info) + 3
    for col, h in enumerate([
        "Lp.", "Kompetencja", "Kategoria", "Definicja",
        "Poziom 1 — Niedostateczny", "Poziom 2 — Podstawowy", "Poziom 3 — Dobry",
        "Poziom 4 — Bardzo dobry", "Poziom 5 — Wybitny",
    ], 1):
        _hdr(ws1, hr, col, h)

    for i, comp in enumerate(profile.competencies, 1):
        r = hr + i
        bg = _LIGHT if i % 2 == 0 else None
        _cell(ws1, r, 1, i, bg=bg)
        _cell(ws1, r, 2, comp.name, bold=True, bg=bg)
        _cell(ws1, r, 3, CATEGORY_LABELS.get(comp.category, comp.category), bg=bg)
        _cell(ws1, r, 4, comp.definition, bg=bg)
        for j, lk in enumerate(LEVEL_KEYS, 5):
            _cell(ws1, r, j, comp.get_level_description(lk), bg=bg)
        ws1.row_dimensions[r].height = 80

    for col, w in enumerate([4, 22, 14, 36, 28, 28, 28, 28, 28], 1):
        ws1.column_dimensions[get_column_letter(col)].width = w

    # ------------------------------------------------------------------ #
    # Arkusz 2: Analiza potrzeb
    # ------------------------------------------------------------------ #
    if assessment:
        ws2 = wb.create_sheet("Analiza_potrzeb")
        ws2.merge_cells("A1:G1")
        t2 = ws2.cell(
            row=1, column=1,
            value=f"Analiza potrzeb — {assessment.assessor_name} ({assessment.assessor_role})"
        )
        t2.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
        t2.fill = PatternFill("solid", fgColor=_NAVY)
        t2.alignment = Alignment(horizontal="center")
        ws2.row_dimensions[1].height = 24

        for col, h in enumerate([
            "Lp.", "Kompetencja", "Aktualny poziom", "Pożądany poziom",
            "Luka", "Ważność (1-5)", "Priorytet",
        ], 1):
            _hdr(ws2, 2, col, h)

        sorted_items = sorted(assessment.items, key=lambda x: x.priority_score, reverse=True)
        for i, item in enumerate(sorted_items, 1):
            r = i + 2
            bg = _LIGHT if i % 2 == 0 else None
            gap_bg = _RED if item.gap >= 2 else (_AMBER if item.gap == 1 else _GREEN)
            _cell(ws2, r, 1, i, bg=bg)
            _cell(ws2, r, 2, item.competency_name, bold=True, bg=bg)
            _cell(ws2, r, 3, f"{item.current_level} — {LEVEL_LABELS[item.current_level]}", bg=bg)
            _cell(ws2, r, 4, f"{item.desired_level} — {LEVEL_LABELS[item.desired_level]}", bg=bg)
            _cell(ws2, r, 5, item.gap, bg=gap_bg)
            _cell(ws2, r, 6, item.importance, bg=bg)
            _cell(ws2, r, 7, item.priority_score, bg=bg)

        for col, w in enumerate([4, 28, 22, 22, 8, 12, 10], 1):
            ws2.column_dimensions[get_column_letter(col)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Profil Kompetencji — Warsztat",
    page_icon="🗺️",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Profil Kompetencji")
    st.caption("Narzędzie warsztatowe — Enterprise Advisors")

    # API key
    default_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not default_key:
        try:
            default_key = st.secrets.get("ANTHROPIC_API_KEY", "")
        except Exception:
            default_key = ""
    if not default_key:
        env_path = os.path.join(os.path.dirname(__file__), ".env")
        if os.path.exists(env_path):
            for line in open(env_path):
                line = line.strip()
                if line.startswith("ANTHROPIC_API_KEY="):
                    default_key = line.split("=", 1)[1].strip().strip('"').strip("'")

    if default_key:
        api_key = default_key
        st.caption("Klucz API: skonfigurowany")
    else:
        api_key = st.text_input("Klucz API Anthropic", type="password")

    st.divider()
    st.markdown(
        "**Jak korzystać:**\n\n"
        "1. **Zakładka 1** — opisz firmę, wybierz kompetencje, wygeneruj opisy\n"
        "2. **Zakładka 2** — oceń poziomy (aktualny vs pożądany)\n"
        "3. **Pobierz pliki** poniżej i wyślij trenerowi\n\n"
        "*Enterprise Advisors*"
    )

    st.divider()
    st.subheader("Pobierz wyniki")

    profile_ready = "profile" in st.session_state

    if profile_ready:
        profile_obj: CompetencyProfile = st.session_state["profile"]
        assessment_obj = st.session_state.get("needs_assessment")

        state_data: dict = {"profile": profile_obj.to_json()}
        if assessment_obj:
            state_data["needs_assessment"] = assessment_obj.to_json()

        safe_name = (
            f"{profile_obj.company.company_name}_{profile_obj.company.position_name}"
            .replace(" ", "_").replace("/", "-")
        )

        st.download_button(
            "Pobierz JSON (dla trenera)",
            data=json.dumps(state_data, ensure_ascii=False, indent=2).encode("utf-8"),
            file_name=f"profil_{safe_name}.json",
            mime="application/json",
            type="primary",
        )

        xlsx_bytes = build_workshop_xlsx(profile_obj, assessment_obj)
        st.download_button(
            "Pobierz XLSX (raport czytelny)",
            data=xlsx_bytes,
            file_name=f"profil_{safe_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        try:
            from stk_export import build_profile_docx
            docx_bytes = build_profile_docx(profile_obj, assessment_obj)
            st.download_button(
                "Pobierz DOCX (raport z wykresem)",
                data=docx_bytes,
                file_name=f"profil_{safe_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        except Exception as _e:
            st.caption(f"DOCX niedostępny: {_e}")

        if not assessment_obj:
            st.caption("Uzupełnij Analizę potrzeb (zakładka 2), aby raporty zawierały lukę kompetencyjną.")
    else:
        st.info("Wygeneruj profil (zakładka 1), aby pobrać wyniki.")


# ---------------------------------------------------------------------------
# Zakładki
# ---------------------------------------------------------------------------
tab1, tab2 = st.tabs(["1. Profil kompetencji", "2. Analiza potrzeb"])


# ===========================================================================
# ZAKŁADKA 1: Profil kompetencji
# ===========================================================================
with tab1:
    st.header("Profiler kompetencji")

    st.subheader("Krok 1: Firma i stanowisko")
    col1, col2 = st.columns(2)
    with col1:
        company_name = st.text_input("Nazwa firmy", key="c_name")
        industry = st.text_input("Branża", key="c_industry",
                                 placeholder="np. IT, produkcja, usługi finansowe")
        size = st.selectbox("Wielkość", ["<50", "50-250", ">250"], key="c_size")
        culture = st.text_area("Kultura / kontekst", key="c_culture", height=80,
                               placeholder="np. flat structure, agile, korporacja...")
    with col2:
        position = st.text_input("Nazwa stanowiska", key="c_position")
        level = st.selectbox("Poziom", [
            "specjalista", "starszy specjalista", "kierownik zespołu",
            "kierownik działu", "dyrektor", "zarząd",
        ], key="c_level")
        tasks = st.text_area("Kluczowe zadania", key="c_tasks", height=100,
                             placeholder="np. zarządzanie zespołem 8 osób, realizacja budżetów...")
        extra = st.text_area("Dodatkowy kontekst (opcj.)", key="c_extra", height=60)

    st.divider()
    st.subheader("Krok 2: Wybierz kompetencje z katalogu (max 10)")

    selected_from_catalog: dict[str, list[str]] = {}
    cat_cols = st.columns(4)
    for idx, (cat_key, cat_label) in enumerate(CATEGORY_LABELS.items()):
        with cat_cols[idx]:
            st.markdown(f"**{cat_label}**")
            selected_from_catalog[cat_key] = st.multiselect(
                f"Wybierz {cat_label.lower()}",
                COMPETENCY_CATALOG[cat_key],
                key=f"cat_select_{cat_key}",
                label_visibility="collapsed",
            )

    all_selected: list[str] = []
    for names in selected_from_catalog.values():
        all_selected.extend(names)

    st.divider()
    col_add1, col_add2 = st.columns([3, 1])
    with col_add1:
        custom_name = st.text_input(
            "Dopisz własną kompetencję (nazwa)",
            key="custom_comp_name",
            placeholder="np. Zarządzanie ryzykiem, Design thinking...",
        )
    with col_add2:
        st.write("")
        st.write("")
        add_custom = st.button("Dodaj do listy", key="add_custom_btn")

    if "custom_competencies" not in st.session_state:
        st.session_state["custom_competencies"] = []

    if add_custom and custom_name.strip():
        if custom_name.strip() not in st.session_state["custom_competencies"]:
            st.session_state["custom_competencies"].append(custom_name.strip())
            st.rerun()

    if st.session_state["custom_competencies"]:
        st.markdown("**Własne kompetencje:**")
        to_remove = []
        for i, cn in enumerate(st.session_state["custom_competencies"]):
            col_c, col_x = st.columns([5, 1])
            col_c.write(f"+ {cn}")
            if col_x.button("Usuń", key=f"rm_custom_{i}"):
                to_remove.append(cn)
        if to_remove:
            for cn in to_remove:
                st.session_state["custom_competencies"].remove(cn)
            st.rerun()

    all_selected.extend(st.session_state.get("custom_competencies", []))
    total_count = len(all_selected)

    if total_count > 10:
        st.warning(f"Wybrano {total_count} kompetencji — max 10. Usuń nadmiarowe.")
    elif total_count > 0:
        st.info(f"Wybrano **{total_count}** kompetencji: {', '.join(all_selected)}")
    else:
        st.info("Wybierz kompetencje z katalogu lub dopisz własne.")

    st.divider()
    st.subheader("Krok 3: Generuj opisy kompetencji")

    if st.button("Generuj profil kompetencji", type="primary",
                 disabled=(not api_key or total_count == 0 or total_count > 10)):
        if not position or not tasks:
            st.warning("Podaj nazwę stanowiska i kluczowe zadania (Krok 1).")
        else:
            company = CompanyProfile(
                company_name=company_name, industry=industry, size=size,
                culture_notes=culture, position_name=position, position_level=level,
                key_tasks=tasks, additional_context=extra,
            )
            with st.spinner(f"Generowanie opisów {total_count} kompetencji (30-90 sek.)..."):
                try:
                    profile = generate_competency_profile(company, all_selected, api_key)
                    st.session_state["profile"] = profile
                    st.session_state["company"] = company
                    st.success(
                        f"Wygenerowano {len(profile.competencies)} kompetencji. "
                        "Przejdź do zakładki 2 → Analiza potrzeb."
                    )
                except Exception as e:
                    st.error(f"Błąd: {e}")

    if "profile" in st.session_state:
        profile: CompetencyProfile = st.session_state["profile"]
        st.divider()
        st.subheader(f"Profil: {profile.company.position_name} ({profile.company.company_name})")
        st.caption(
            f"{len(profile.competencies)} kompetencji"
            + (f" | {profile.created_at[:19]}" if profile.created_at else "")
        )

        comp_to_remove = None
        for i, comp in enumerate(profile.competencies):
            cat_display = CATEGORY_LABELS.get(comp.category, comp.category)
            with st.expander(f"**{i+1}. {comp.name}** [{cat_display}]", expanded=(i < 2)):
                st.markdown(f"**Definicja:** {comp.definition}")
                if comp.indicators:
                    st.markdown("**Wskaźniki behawioralne:**")
                    for ind in comp.indicators:
                        st.markdown(f"- {ind}")
                level_cols = st.columns(5)
                for j, lk in enumerate(LEVEL_KEYS):
                    with level_cols[j]:
                        desc = comp.get_level_description(lk)
                        color = LEVEL_COLORS[lk]
                        st.markdown(
                            f"<div style='background:{color}; color:white; padding:8px; "
                            f"border-radius:6px; font-size:0.8em; min-height:120px;'>"
                            f"<b>{lk}. {LEVEL_LABELS[lk]}</b><br><br>{desc}</div>",
                            unsafe_allow_html=True,
                        )
                if st.button(f"Usuń kompetencję '{comp.name}'", key=f"rm_comp_{i}"):
                    comp_to_remove = i

        if comp_to_remove is not None:
            profile.competencies.pop(comp_to_remove)
            st.session_state["profile"] = profile
            st.rerun()

        st.divider()
        col_d1, col_d2 = st.columns([3, 1])
        with col_d1:
            new_comp_name = st.text_input(
                "Dopisz kompetencję do profilu (AI wygeneruje opis)",
                key="add_comp_to_profile",
                placeholder="Wpisz nazwę kompetencji...",
            )
        with col_d2:
            st.write("")
            st.write("")
            add_to_profile = st.button("Dopisz i generuj opis", key="add_comp_btn")

        if add_to_profile and new_comp_name.strip() and api_key:
            company = st.session_state.get("company", profile.company)
            with st.spinner(f"Generowanie opisu '{new_comp_name.strip()}'..."):
                try:
                    new_comp = generate_single_competency(new_comp_name.strip(), company, api_key)
                    profile.competencies.append(new_comp)
                    st.session_state["profile"] = profile
                    st.success(f"Dodano: {new_comp.name}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Błąd: {e}")


# ===========================================================================
# ZAKŁADKA 2: Analiza potrzeb
# ===========================================================================
with tab2:
    st.header("Analiza potrzeb szkoleniowych")

    if "profile" not in st.session_state:
        st.info("Najpierw wygeneruj profil kompetencji w zakładce 1.")
        st.stop()

    profile: CompetencyProfile = st.session_state["profile"]
    st.markdown(
        "Oceń **aktualny** i **pożądany** poziom każdej kompetencji (1-5). "
        "System wyliczy lukę kompetencyjną i priorytet szkoleniowy."
    )

    assessor_name = st.text_input("Imię i nazwisko oceniającego", key="assess_name")
    assessor_role = st.selectbox(
        "Rola", ["przełożony", "uczestnik (samoocena)", "trener"], key="assess_role"
    )
    st.divider()

    needs_items: list[NeedsAssessmentItem] = []
    for i, comp in enumerate(profile.competencies):
        col1, col2, col3, col4 = st.columns([3, 1.5, 1.5, 1])
        with col1:
            st.markdown(f"**{comp.name}**")
        with col2:
            current = st.select_slider(
                "Aktualny", options=[1, 2, 3, 4, 5], value=3,
                format_func=lambda x: f"{x} — {LEVEL_LABELS[x][:12]}",
                key=f"curr_{i}",
            )
        with col3:
            desired = st.select_slider(
                "Pożądany", options=[1, 2, 3, 4, 5], value=4,
                format_func=lambda x: f"{x} — {LEVEL_LABELS[x][:12]}",
                key=f"des_{i}",
            )
        with col4:
            importance = st.select_slider("Waga", options=[1, 2, 3, 4, 5], value=3, key=f"imp_{i}")
        needs_items.append(NeedsAssessmentItem(
            competency_name=comp.name, current_level=current,
            desired_level=desired, importance=importance, assessor=assessor_role,
        ))

    if st.button("Zapisz ocenę potrzeb", type="primary"):
        assessment = NeedsAssessment(
            items=needs_items, assessor_name=assessor_name, assessor_role=assessor_role,
        )
        st.session_state["needs_assessment"] = assessment
        st.success("Ocena zapisana. Pobierz pliki z panelu bocznego i wyślij trenerowi.")

    if "needs_assessment" in st.session_state:
        assessment: NeedsAssessment = st.session_state["needs_assessment"]
        st.divider()
        st.subheader("Wyniki analizy potrzeb")

        sorted_items = sorted(assessment.items, key=lambda x: x.priority_score, reverse=True)
        header_cols = st.columns([3, 1, 1, 1, 1])
        for col_w, label in zip(header_cols, ["Kompetencja", "Aktualny", "Pożądany", "Luka", "Priorytet"]):
            col_w.markdown(f"**{label}**")

        for item in sorted_items:
            cols = st.columns([3, 1, 1, 1, 1])
            cols[0].write(item.competency_name)
            cols[1].write(f"{item.current_level} — {LEVEL_LABELS[item.current_level][:10]}")
            cols[2].write(f"{item.desired_level} — {LEVEL_LABELS[item.desired_level][:10]}")
            gap_color = "red" if item.gap >= 2 else ("orange" if item.gap == 1 else "green")
            cols[3].markdown(f":{gap_color}[**{item.gap}**]")
            cols[4].write(str(item.priority_score))

        st.divider()
        st.info(
            "Gotowe! Pobierz pliki z panelu bocznego po lewej:\n\n"
            "- **JSON** — plik dla trenera (import do Asystenta STK)\n"
            "- **XLSX** — czytelny raport z opisami kompetencji i luką"
        )
