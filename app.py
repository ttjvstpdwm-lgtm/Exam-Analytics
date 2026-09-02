from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import html
import re
import sys
import unicodedata

import altair as alt
import pandas as pd
import streamlit as st

from exam_tools import (
    apply_corrections,
    build_corrected_workbook,
    excel_source_from_bytes,
    filter_by_class,
    find_exam_sheets,
    load_corrections,
    load_exam,
    normalize_label,
    question_summary,
    save_student_corrections,
    student_totals,
)


APP_DIR = Path(__file__).resolve().parent
VENDOR_DIR = APP_DIR / "vendor"
if VENDOR_DIR.exists() and str(VENDOR_DIR) not in sys.path:
    sys.path.insert(0, str(VENDOR_DIR))

CORRECTIONS_PATH = APP_DIR / "data" / "manual_scores.csv"
DEFAULT_EXCEL_PATH = Path(
    "/Users/chizzoli/Library/CloudStorage/OneDrive-UniversitàCommercialeLuigiBocconi/30705 - Marketing/Esami/Giugno 2026/30705_2026_06_22_ITA_FREQ.xlsx"
)
OFFICIAL_RESULTS_PATH = Path(
    "/Users/chizzoli/Library/CloudStorage/OneDrive-UniversitàCommercialeLuigiBocconi/30705 - Marketing/Esami/Giugno 2026/Esiti_20260622_30705_TR01_S.xls"
)


st.set_page_config(
    page_title="Analisi esame Blackboard",
    page_icon="BB",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.4rem; }
    div[data-testid="stMetric"] {
        border: 1px solid #e7e2d8;
        border-radius: 8px;
        padding: 12px 14px;
        background: #fffdfa;
    }
    .status-pill {
        display: inline-block;
        padding: 3px 8px;
        border-radius: 999px;
        font-size: 0.82rem;
        font-weight: 650;
        border: 1px solid transparent;
    }
    .status-ok { background: #e9f7ef; color: #11623b; border-color: #bfe6d0; }
    .status-bad { background: #fdebea; color: #9b1c1c; border-color: #f5c2c0; }
    .status-partial { background: #fff6df; color: #7a4a00; border-color: #f3d390; }
    .status-missing { background: #eef1f4; color: #39424e; border-color: #d5dbe1; }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def cached_load_exam(
    path_text: str,
    uploaded_bytes: bytes | None,
    uploaded_name: str | None,
    path_mtime: float | None,
    selected_sheets: tuple[str, ...],
):
    if uploaded_bytes:
        return load_exam(excel_source_from_bytes(uploaded_bytes), uploaded_name or "File caricato", selected_sheets)
    return load_exam(Path(path_text), Path(path_text).name, selected_sheets)


@st.cache_data(show_spinner=False)
def cached_find_exam_sheets(
    path_text: str,
    uploaded_bytes: bytes | None,
    uploaded_name: str | None,
    path_mtime: float | None,
) -> list[str]:
    if uploaded_bytes:
        return find_exam_sheets(excel_source_from_bytes(uploaded_bytes))
    return find_exam_sheets(Path(path_text))


def source_for_export(path_text: str, uploaded_bytes: bytes | None) -> str | Path | bytes:
    if uploaded_bytes:
        return uploaded_bytes
    return Path(path_text)


def fmt_num(value: float | int | None, digits: int = 2) -> str:
    if value is None or pd.isna(value):
        return "-"
    return f"{float(value):.{digits}f}"


def plain_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def safe_filename(value: object) -> str:
    text = plain_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_") or "studente"


def normalize_person_name(value: object) -> str:
    text = unicodedata.normalize("NFKD", plain_text(value))
    text = text.encode("ascii", "ignore").decode("ascii").upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text).strip()
    return re.sub(r"\s+", " ", text)


def truthy_vision_value(value: object) -> bool:
    text = plain_text(value).strip().lower()
    if not text:
        return False
    return text in {
        "1",
        "true",
        "vero",
        "yes",
        "y",
        "si",
        "sì",
        "x",
        "ok",
        "iscritto",
        "iscritta",
        "prenotato",
        "prenotata",
        "presente",
    }


def roster_ids_from_upload(uploaded_roster) -> set[str]:
    if uploaded_roster is None:
        return set()
    content = uploaded_roster.getvalue()
    if uploaded_roster.name.lower().endswith(".csv"):
        roster_df = pd.read_csv(BytesIO(content))
    else:
        roster_df = pd.read_excel(BytesIO(content))
    if roster_df.empty:
        return set()

    candidates = {
        "username",
        "user name",
        "matricola",
        "id",
        "student id",
        "student_id",
        "user id",
        "userid",
    }
    selected_col = None
    for column in roster_df.columns:
        if str(column).strip().lower() in candidates:
            selected_col = column
            break
    if selected_col is None:
        selected_col = roster_df.columns[0]

    return {
        normalize_label(value)
        for value in roster_df[selected_col].dropna().tolist()
        if normalize_label(value)
    }


def roster_students_from_upload(uploaded_roster, candidate_students: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, object]]:
    if uploaded_roster is None:
        return candidate_students.iloc[0:0].copy(), {"rows": 0, "matched": 0, "unmatched": [], "matched_by": {}}

    content = uploaded_roster.getvalue()
    if uploaded_roster.name.lower().endswith(".csv"):
        roster_df = pd.read_csv(BytesIO(content))
    else:
        roster_df = pd.read_excel(BytesIO(content))

    candidate_students = candidate_students.copy()
    candidate_students["__username_key"] = candidate_students["username"].map(normalize_label)
    candidate_students["__full_name_key"] = candidate_students["full_name"].map(normalize_person_name)
    candidate_students["__first_last_key"] = (
        candidate_students["first_name"].map(plain_text) + " " + candidate_students["last_name"].map(plain_text)
    ).map(normalize_person_name)

    by_username = {
        row["__username_key"]: row["student_uid"]
        for _, row in candidate_students.iterrows()
        if row["__username_key"]
    }
    by_name: dict[str, set[str]] = {}
    for _, row in candidate_students.iterrows():
        for key in [row["__full_name_key"], row["__first_last_key"]]:
            if key:
                by_name.setdefault(key, set()).add(row["student_uid"])

    id_columns = {"username", "user name", "matricola", "student id", "student_id", "user id", "userid", "id"}
    matched_uids: set[str] = set()
    matched_by = {"matricola": 0, "nome": 0}
    unmatched: list[str] = []

    for _, row in roster_df.iterrows():
        row_matched = False
        for column in roster_df.columns:
            if str(column).strip().lower() in id_columns:
                key = normalize_label(row.get(column))
                if key in by_username:
                    matched_uids.add(by_username[key])
                    matched_by["matricola"] += 1
                    row_matched = True
                    break

        if not row_matched:
            name_candidates = [
                row.get("Nome"),
                row.get("Full Name"),
                row.get("Nome completo"),
                f"{plain_text(row.get('Nome2'))} {plain_text(row.get('Cognome'))}",
                f"{plain_text(row.get('First Name'))} {plain_text(row.get('Last Name'))}",
            ]
            for name in name_candidates:
                key = normalize_person_name(name)
                matches = list(by_name.get(key, set()))
                if len(matches) == 1:
                    matched_uids.add(matches[0])
                    matched_by["nome"] += 1
                    row_matched = True
                    break

        if not row_matched:
            label = plain_text(row.get("Nome")) or plain_text(row.get("Nome2")) or plain_text(row.iloc[0])
            if label:
                unmatched.append(label)

    selected = (
        candidate_students[candidate_students["student_uid"].isin(matched_uids)]
        .drop(columns=["__username_key", "__full_name_key", "__first_last_key"], errors="ignore")
        .copy()
    )
    stats = {
        "rows": len(roster_df),
        "matched": len(selected),
        "unmatched": unmatched,
        "matched_by": matched_by,
    }
    return selected, stats


GRADE_BANDS = [
    "Insuff. gravi (<15)",
    "Insuff. lievi (15-17)",
    "18-20",
    "21-23",
    "24-27",
    "28-29",
    "30-31",
]


def grade_distribution(totals_df: pd.DataFrame, score_column: str) -> pd.DataFrame:
    dist_data = totals_df.dropna(subset=[score_column]).copy()
    dist_data["fascia"] = pd.cut(
        dist_data[score_column],
        bins=[float("-inf"), 15, 18, 21, 24, 28, 30, 32],
        labels=GRADE_BANDS,
        right=False,
    )
    counts = (
        dist_data["fascia"]
        .value_counts(sort=False)
        .reindex(GRADE_BANDS, fill_value=0)
        .rename_axis("fascia")
        .reset_index(name="studenti")
    )
    total_students = max(1, len(dist_data))
    counts["percentuale"] = counts["studenti"] / total_students
    counts["etichetta"] = counts.apply(lambda row: f"{int(row['studenti'])} ({row['percentuale']:.0%})", axis=1)
    return counts


def score_column(totals_df: pd.DataFrame, preferred: str, fallback: str) -> str:
    if preferred in totals_df.columns and totals_df[preferred].notna().any():
        return preferred
    return fallback


def distribution_chart(dist_counts: pd.DataFrame, x_title: str):
    dist_labels = dist_counts[dist_counts["studenti"] > 0].copy()
    bars = (
        alt.Chart(dist_counts)
        .mark_bar(color="#4f7cac")
        .encode(
            x=alt.X("fascia:N", title=x_title, sort=GRADE_BANDS),
            y=alt.Y("studenti:Q", title="Studenti"),
            tooltip=[
                alt.Tooltip("fascia:N", title="Fascia"),
                alt.Tooltip("studenti:Q", title="Studenti"),
                alt.Tooltip("percentuale:Q", title="Percentuale", format=".1%"),
            ],
        )
    )
    labels = (
        alt.Chart(dist_labels)
        .mark_text(dy=-8, color="#27313d", fontWeight="bold", fontSize=12)
        .encode(
            x=alt.X("fascia:N", sort=GRADE_BANDS),
            y=alt.Y("studenti:Q"),
            text="etichetta:N",
        )
    )
    return (bars + labels).properties(height=310).configure_axis(labelLimit=140)


def status_html(status: str) -> str:
    css = {
        "Corretta": "status-ok",
        "Errata": "status-bad",
        "Parziale": "status-partial",
        "Da correggere": "status-missing",
        "Da verificare": "status-missing",
    }.get(status, "status-missing")
    return f'<span class="status-pill {css}">{status}</span>'


def score_validation_errors(edited: pd.DataFrame) -> list[str]:
    errors: list[str] = []
    for _, row in edited.iterrows():
        score = row.get("manual_score")
        possible = row.get("possible_points")
        if pd.isna(score):
            continue
        if score < 0:
            errors.append(f"Domanda {int(row['question_num'])}: il manual score non puo essere negativo.")
        if not pd.isna(possible) and score > possible:
            errors.append(
                f"Domanda {int(row['question_num'])}: il manual score ({score:g}) supera i punti possibili ({possible:g})."
            )
    return errors


def docx_escape(value: object) -> str:
    return html.escape(plain_text(value), quote=False)


def paragraph_xml(text: object = "", style: str | None = None) -> str:
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    escaped = docx_escape(text)
    if not escaped:
        return f"<w:p>{style_xml}</w:p>"
    return f'<w:p>{style_xml}<w:r><w:t xml:space="preserve">{escaped}</w:t></w:r></w:p>'


def bullet_xml(text: object) -> str:
    escaped = docx_escape(text)
    return (
        "<w:p>"
        '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{escaped}</w:t></w:r>'
        "</w:p>"
    )


def document_table_xml(rows: list[list[object]]) -> str:
    grid_cols = "".join('<w:gridCol w:w="2600"/>' for _ in rows[0])
    xml = [
        "<w:tbl>",
        '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>',
        f"<w:tblGrid>{grid_cols}</w:tblGrid>",
    ]
    for row in rows:
        xml.append("<w:tr>")
        for cell in row:
            xml.append(f"<w:tc><w:tcPr><w:tcW w:w=\"2600\" w:type=\"dxa\"/></w:tcPr>{paragraph_xml(cell)}</w:tc>")
        xml.append("</w:tr>")
    xml.append("</w:tbl>")
    return "".join(xml)


def build_review_docx(
    selected_total: pd.Series,
    student_questions: pd.DataFrame,
    include_question_text: bool,
    include_correct_answer: bool,
) -> bytes:
    review_questions = student_questions[student_questions["status"].isin(["Errata", "Parziale"])].sort_values("question_num")
    title = f"Scheda visione compiti - {plain_text(selected_total.get('full_name'))}"
    body: list[str] = [
        paragraph_xml(title, "Title"),
        paragraph_xml("Riepilogo studente", "Heading1"),
        document_table_xml(
            [
                ["Progetto", "Esame", "Finale rounded"],
                [
                    fmt_num(selected_total.get("progetto")),
                    fmt_num(selected_total.get("totale_esame", selected_total.get("totale_integrato"))),
                    fmt_num(selected_total.get("finale_rounded")),
                ],
            ]
        ),
        paragraph_xml("Domande errate o parzialmente corrette", "Heading1"),
    ]

    if review_questions.empty:
        body.append(paragraph_xml("Non risultano domande errate o parzialmente corrette."))
    else:
        for _, row in review_questions.iterrows():
            heading = (
                f"Domanda {int(row['question_num'])}"
                f" - ID {plain_text(row.get('question_id')) or '-'}"
                f" - {plain_text(row.get('status'))}"
                f" ({fmt_num(row.get('integrated_score'))}/{fmt_num(row.get('possible_points'))})"
            )
            body.append(paragraph_xml(heading, "Heading2"))
            body.append(bullet_xml(f"Risposta dello studente: {plain_text(row.get('answer')) or '-'}"))
            if include_question_text:
                body.append(bullet_xml(f"Testo domanda: {plain_text(row.get('question_text')) or '-'}"))
            if include_correct_answer:
                correct_answer = plain_text(row.get("correct_answer")) or "Non disponibile nel file"
                body.append(bullet_xml(f"Risposta corretta: {correct_answer}"))
            if int(row["question_num"]) in {27, 28, 29} and plain_text(row.get("comment")):
                body.append(bullet_xml(f"Commento correzione: {plain_text(row.get('comment'))}"))

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body>"
        + "".join(body)
        + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
        "</w:body></w:document>"
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:spacing w:before="360" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="80"/></w:pPr><w:rPr><w:b/><w:sz w:val="21"/></w:rPr></w:style>'
        '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr></w:style>'
        "</w:styles>"
    )
    numbering_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        "</w:numbering>"
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
        "</Types>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )
    document_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
        "</Relationships>"
    )

    buffer = BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", rels_xml)
        docx.writestr("word/_rels/document.xml.rels", document_rels_xml)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/styles.xml", styles_xml)
        docx.writestr("word/numbering.xml", numbering_xml)
    return buffer.getvalue()


def unique_zip_name(base_name: str, used_names: set[str]) -> str:
    stem = base_name.removesuffix(".docx")
    name = base_name
    counter = 2
    while name in used_names:
        name = f"{stem}_{counter}.docx"
        counter += 1
    used_names.add(name)
    return name


def build_review_zip(
    selected_students: pd.DataFrame,
    all_scored: pd.DataFrame,
    include_question_text: bool,
    include_correct_answer: bool,
) -> bytes:
    buffer = BytesIO()
    used_names: set[str] = set()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
        for _, student in selected_students.sort_values(["classe", "last_name", "first_name", "username"], kind="stable").iterrows():
            student_questions = all_scored[all_scored["student_uid"].eq(student["student_uid"])].sort_values("question_num")
            file_base = (
                f"scheda_visione_compiti_"
                f"classe_{safe_filename(student.get('classe'))}_"
                f"{safe_filename(student.get('full_name') or student.get('username'))}.docx"
            )
            file_name = unique_zip_name(file_base, used_names)
            archive.writestr(
                file_name,
                build_review_docx(
                    student,
                    student_questions,
                    include_question_text=include_question_text,
                    include_correct_answer=include_correct_answer,
                ),
            )
    return buffer.getvalue()


def page_break_xml() -> str:
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def extract_docx_body_content(docx_bytes: bytes) -> str:
    with ZipFile(BytesIO(docx_bytes), "r") as docx:
        document_xml = docx.read("word/document.xml").decode("utf-8")
    body_start = document_xml.index("<w:body>") + len("<w:body>")
    section_start = document_xml.rindex("<w:sectPr")
    return document_xml[body_start:section_start]


def replace_docx_body(template_docx_bytes: bytes, body_content: str) -> bytes:
    section_xml = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{body_content}{section_xml}</w:body></w:document>"
    )
    output = BytesIO()
    with ZipFile(BytesIO(template_docx_bytes), "r") as source, ZipFile(output, "w", ZIP_DEFLATED) as target:
        for item in source.infolist():
            content = document_xml.encode("utf-8") if item.filename == "word/document.xml" else source.read(item.filename)
            target.writestr(item, content)
    return output.getvalue()


def build_combined_review_docx(
    selected_students: pd.DataFrame,
    all_scored: pd.DataFrame,
    include_question_text: bool,
    include_correct_answer: bool,
) -> bytes:
    fragments: list[str] = []
    template_docx: bytes | None = None
    sorted_students = selected_students.sort_values(["classe", "last_name", "first_name", "username"], kind="stable")
    for index, (_, student) in enumerate(sorted_students.iterrows()):
        student_questions = all_scored[all_scored["student_uid"].eq(student["student_uid"])].sort_values("question_num")
        student_docx = build_review_docx(
            student,
            student_questions,
            include_question_text=include_question_text,
            include_correct_answer=include_correct_answer,
        )
        if template_docx is None:
            template_docx = student_docx
        if index > 0:
            fragments.append(page_break_xml())
        fragments.append(extract_docx_body_content(student_docx))
    if template_docx is None:
        return build_review_docx(pd.Series(dtype=object), pd.DataFrame(), include_question_text, include_correct_answer)
    return replace_docx_body(template_docx, "".join(fragments))


def excel_value(value: object) -> int | float | None:
    if value is None or pd.isna(value):
        return None
    number = float(value)
    if number.is_integer():
        return int(number)
    return round(number, 2)


def find_header_column(sheet, header_row: int, names: list[str], fallback: int) -> int:
    normalized_names = [name.strip().lower() for name in names]
    for col in range(sheet.ncols):
        value = str(sheet.cell_value(header_row, col)).strip().lower()
        if any(name in value for name in normalized_names):
            return col
    return fallback


def compile_official_results_xls(template_source: str | Path | bytes, totals_df: pd.DataFrame) -> tuple[bytes, dict[str, object]]:
    import xlrd
    from xlutils.copy import copy as copy_workbook

    if isinstance(template_source, bytes):
        read_book = xlrd.open_workbook(file_contents=template_source, formatting_info=True)
    else:
        read_book = xlrd.open_workbook(str(template_source), formatting_info=True)

    read_sheet = read_book.sheet_by_index(0)
    header_row = None
    for row in range(min(read_sheet.nrows, 20)):
        row_values = [str(read_sheet.cell_value(row, col)).strip().lower() for col in range(read_sheet.ncols)]
        if "matricola" in row_values and "voto" in row_values:
            header_row = row
            break
    if header_row is None:
        header_row = 4

    id_col = find_header_column(read_sheet, header_row, ["matricola", "id"], 0)
    voto_col = find_header_column(read_sheet, header_row, ["voto"], 16)
    esame_col = find_header_column(read_sheet, header_row, ["esame scritto"], 22)
    progetto_col = find_header_column(read_sheet, header_row, ["progetto"], 23)

    score_map = {
        normalize_label(row["username"]): row
        for _, row in totals_df.iterrows()
        if normalize_label(row.get("username"))
    }

    write_book = copy_workbook(read_book)
    write_sheet = write_book.get_sheet(0)
    matched_ids: set[str] = set()
    written_rows = 0

    for row_idx in range(header_row + 1, read_sheet.nrows):
        student_id = normalize_label(read_sheet.cell_value(row_idx, id_col))
        if student_id not in score_map:
            continue
        row = score_map[student_id]
        values = [
            (voto_col, excel_value(row.get("finale_rounded"))),
            (esame_col, excel_value(row.get("totale_esame"))),
            (progetto_col, excel_value(row.get("progetto"))),
        ]
        for col_idx, value in values:
            if value is not None:
                write_sheet.write(row_idx, col_idx, value)
        matched_ids.add(student_id)
        written_rows += 1

    output = BytesIO()
    write_book.save(output)
    missing_in_template = sorted(set(score_map) - matched_ids)
    stats = {
        "matched": len(matched_ids),
        "written_rows": written_rows,
        "students_in_exam": len(score_map),
        "missing_in_template": missing_in_template,
        "id_column": id_col + 1,
        "voto_column": voto_col + 1,
        "esame_column": esame_col + 1,
        "progetto_column": progetto_col + 1,
    }
    return output.getvalue(), stats


st.title("Analisi esame Blackboard")

with st.sidebar:
    st.header("Origine dati")
    uploaded = st.file_uploader("Carica un file Excel", type=["xlsx"])
    path_text = st.text_input("Oppure usa questo percorso locale", value=str(DEFAULT_EXCEL_PATH))
    st.caption("I fogli con colonne Question ID vengono rilevati automaticamente.")

uploaded_bytes = uploaded.getvalue() if uploaded else None
uploaded_name = uploaded.name if uploaded else None
local_path = Path(path_text).expanduser()
path_mtime = None
if not uploaded_bytes:
    try:
        path_mtime = local_path.stat().st_mtime
    except OSError:
        path_mtime = None

if not uploaded_bytes and not local_path.exists():
    st.info("Carica un file Excel dalla sidebar oppure inserisci un percorso locale valido per iniziare.")
    st.stop()

try:
    available_exam_sheets = cached_find_exam_sheets(path_text, uploaded_bytes, uploaded_name, path_mtime)
except Exception as exc:
    st.error(f"Non riesco a leggere i fogli del file: {exc}")
    st.stop()

if not available_exam_sheets:
    st.error("Non ho trovato fogli con colonne 'Question ID n'.")
    st.stop()

source_signature = (
    f"upload:{uploaded_name}:{len(uploaded_bytes)}"
    if uploaded_bytes
    else f"path:{local_path}:{path_mtime}"
)
if st.session_state.get("exam_sheet_source_signature") != source_signature:
    st.session_state["exam_sheet_selection"] = available_exam_sheets
    st.session_state["exam_sheet_source_signature"] = source_signature

with st.sidebar:
    selected_exam_sheets = st.multiselect(
        "Fogli esame da includere",
        options=available_exam_sheets,
        key="exam_sheet_selection",
        help="Puoi includere tutti i fogli o selezionare solo attending/non attending.",
    )
    st.caption(f"Fogli esame rilevati: {', '.join(available_exam_sheets)}")

if not selected_exam_sheets:
    st.warning("Seleziona almeno un foglio esame dalla sidebar.")
    st.stop()

try:
    exam = cached_load_exam(path_text, uploaded_bytes, uploaded_name, path_mtime, tuple(selected_exam_sheets))
except Exception as exc:
    st.error(f"Non riesco a leggere il file: {exc}")
    st.stop()

corrections = load_corrections(CORRECTIONS_PATH)
scored = apply_corrections(exam.long_df, corrections)
totals = student_totals(scored)

classes = ["Tutte"] + sorted([c for c in scored["classe"].dropna().astype(str).unique() if c])
with st.sidebar:
    st.header("Filtri")
    selected_class = st.selectbox("Classe", classes)
    st.caption(f"Fogli analizzati: {', '.join(exam.exam_sheets)}")
    st.caption(f"Correzioni salvate: {len(corrections)}")

filtered = filter_by_class(scored, selected_class)
filtered_totals = student_totals(filtered)
summary = question_summary(filtered)

tab_dashboard, tab_students, tab_export = st.tabs(["Riepilogo", "Singolo esame", "Export"])

with tab_dashboard:
    st.subheader("Riepilogo per classe")
    if filtered_totals.empty:
        st.info("Nessuno studente nel filtro selezionato.")
    else:
        exam_score_col = score_column(filtered_totals, "totale_esame", "totale_integrato")
        final_score_available = "finale_rounded" in filtered_totals.columns and filtered_totals["finale_rounded"].notna().any()

        st.markdown("### Voti esame")
        metric_cols = st.columns(5)
        metric_cols[0].metric("Studenti", f"{filtered_totals['student_uid'].nunique():,}".replace(",", "."))
        metric_cols[1].metric("Media totale esame", fmt_num(filtered_totals[exam_score_col].mean()))
        metric_cols[2].metric("Mediana totale esame", fmt_num(filtered_totals[exam_score_col].median()))
        metric_cols[3].metric("Con voto esame", int(filtered_totals[exam_score_col].notna().sum()))
        metric_cols[4].metric("Da correggere", int(filtered["is_missing_score"].sum()))

        exam_dist_counts = grade_distribution(filtered_totals, exam_score_col)

        left, right = st.columns([1.05, 1])
        with left:
            st.altair_chart(distribution_chart(exam_dist_counts, "Fascia voto esame"), width="stretch")
        with right:
            q_chart_data = summary.copy()
            q_chart_data["score_pct_medio"] = q_chart_data["score_pct_medio"].fillna(0)
            q_chart = (
                alt.Chart(q_chart_data)
                .mark_bar(color="#7b9e87")
                .encode(
                    x=alt.X("question_num:O", title="Domanda"),
                    y=alt.Y("score_pct_medio:Q", title="Punteggio medio / punti", axis=alt.Axis(format="%")),
                    tooltip=[
                        alt.Tooltip("question_num:O", title="Domanda"),
                        alt.Tooltip("score_pct_medio:Q", title="Media", format=".1%"),
                        alt.Tooltip("da_correggere:Q", title="Da correggere"),
                    ],
                )
                .properties(height=310)
            )
            st.altair_chart(q_chart, width="stretch")

        st.markdown("### Voto finale rounded")
        if final_score_available:
            final_metric_cols = st.columns(3)
            final_metric_cols[0].metric("Studenti con voto finale", int(filtered_totals["finale_rounded"].notna().sum()))
            final_metric_cols[1].metric("Media finale rounded", fmt_num(filtered_totals["finale_rounded"].mean()))
            final_metric_cols[2].metric("Mediana finale rounded", fmt_num(filtered_totals["finale_rounded"].median()))
            final_dist_counts = grade_distribution(filtered_totals, "finale_rounded")
            st.altair_chart(distribution_chart(final_dist_counts, "Fascia voto finale rounded"), width="stretch")
        else:
            st.info("Nel filtro selezionato non ci sono valori in FINALE ROUNDED.")

        st.markdown("### Schede visione compiti")
        bulk_cols = st.columns([1, 1, 2])
        bulk_include_question_text = bulk_cols[0].checkbox("Includi testo domanda", value=False, key="bulk_question_text")
        bulk_include_correct_answer = bulk_cols[1].checkbox("Includi risposta corretta", value=False, key="bulk_correct_answer")
        selection_mode = bulk_cols[2].radio(
            "Studenti da includere",
            ["Tutti nel filtro classe", "Solo iscritti alla visione"],
            horizontal=True,
        )
        output_format = st.radio(
            "Formato output",
            ["Unico file Word", "ZIP con file singoli"],
            horizontal=True,
        )

        main_registered_mask = filtered_totals["visione_compiti"].map(truthy_vision_value)
        main_registered = filtered_totals[main_registered_mask].copy()
        roster_students = filtered_totals.iloc[0:0].copy()
        roster_stats = {"rows": 0, "matched": 0, "unmatched": [], "matched_by": {}}

        if selection_mode == "Solo iscritti alla visione":
            uploaded_roster = st.file_uploader(
                "Elenco iscritti alla visione compiti",
                type=["xlsx", "csv"],
                key="vision_roster_upload",
                help="Puoi caricare il file Microsoft Forms: uso Matricola/Username se presente, altrimenti Nome o Nome2+Cognome.",
            )
            if uploaded_roster is not None:
                try:
                    roster_students, roster_stats = roster_students_from_upload(uploaded_roster, filtered_totals)
                except Exception as exc:
                    st.error(f"Non riesco a leggere l'elenco iscritti: {exc}")
                    roster_students = filtered_totals.iloc[0:0].copy()

            selected_students = (
                pd.concat([main_registered, roster_students], ignore_index=True)
                .drop_duplicates("student_uid")
                .sort_values(["classe", "last_name", "first_name", "username"], kind="stable")
            )
            st.caption(
                f"Iscritti da colonna nel file principale: {len(main_registered)} | "
                f"Iscritti da elenco caricato trovati nel filtro: {len(roster_students)}"
            )
            if uploaded_roster is not None:
                matched_by = roster_stats.get("matched_by", {})
                st.caption(
                    f"Righe lette dall'elenco: {roster_stats.get('rows', 0)} | "
                    f"Match per matricola: {matched_by.get('matricola', 0)} | "
                    f"Match per nome: {matched_by.get('nome', 0)}"
                )
                if roster_stats.get("unmatched"):
                    st.warning(
                        "Iscritti non trovati nel filtro corrente: "
                        + ", ".join(roster_stats["unmatched"][:8])
                        + ("..." if len(roster_stats["unmatched"]) > 8 else "")
                    )
        else:
            selected_students = filtered_totals.copy()

        st.caption(f"Schede che verranno generate: {len(selected_students)}")
        if selected_students.empty:
            st.warning("Nessuno studente selezionato per la generazione delle schede.")
        else:
            if output_format == "Unico file Word":
                combined_docx = build_combined_review_docx(
                    selected_students,
                    scored,
                    include_question_text=bulk_include_question_text,
                    include_correct_answer=bulk_include_correct_answer,
                )
                st.download_button(
                    "Scarica unico Word schede visione compiti",
                    data=combined_docx,
                    file_name=f"schede_visione_compiti_{safe_filename(selected_class)}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                )
            else:
                bulk_zip = build_review_zip(
                    selected_students,
                    scored,
                    include_question_text=bulk_include_question_text,
                    include_correct_answer=bulk_include_correct_answer,
                )
                st.download_button(
                    "Scarica ZIP schede visione compiti",
                    data=bulk_zip,
                    file_name=f"schede_visione_compiti_{safe_filename(selected_class)}.zip",
                    mime="application/zip",
                    type="primary",
                )

        st.subheader("Analisi domande")
        q_table = summary[
            [
                "question_id",
                "punti_possibili",
                "media_autoscore",
                "media_integrata",
                "da_correggere",
                "parziali",
                "errate",
                "testo",
            ]
        ].rename(
            columns={
                "question_id": "Question ID",
                "punti_possibili": "Punti",
                "media_autoscore": "Media autoscore",
                "media_integrata": "Media integrata",
                "da_correggere": "Da correggere",
                "parziali": "Parziali",
                "errate": "Errate",
                "testo": "Testo",
            }
        )
        st.dataframe(
            q_table,
            width="stretch",
            hide_index=True,
            column_config={
                "Testo": st.column_config.TextColumn("Testo", width="large"),
            },
        )

        if selected_class == "Tutte":
            st.subheader("Confronto classi")
            class_totals = totals.groupby("classe", dropna=False).agg(
                studenti=("student_uid", "nunique"),
                media_totale_esame=("totale_esame", "mean"),
                media_finale_rounded=("finale_rounded", "mean"),
                mediana_totale_esame=("totale_esame", "median"),
                da_correggere=("domande_da_correggere", "sum"),
            ).reset_index()
            st.dataframe(
                class_totals.rename(
                    columns={
                        "classe": "Classe",
                        "studenti": "Studenti",
                        "media_totale_esame": "Media totale esame",
                        "media_finale_rounded": "Media finale rounded",
                        "mediana_totale_esame": "Mediana totale esame",
                        "da_correggere": "Da correggere",
                    }
                ),
                width="stretch",
                hide_index=True,
            )

with tab_students:
    st.subheader("Singolo esame")
    student_pool = filtered_totals.copy()
    if student_pool.empty:
        st.info("Nessuno studente nel filtro selezionato.")
    else:
        selected_student_label = st.selectbox(
            "Studente",
            student_pool["student_label"].tolist(),
            index=0,
        )
        selected_uid = student_pool.loc[student_pool["student_label"] == selected_student_label, "student_uid"].iloc[0]
        student_questions = scored[scored["student_uid"] == selected_uid].sort_values("question_num").copy()
        selected_total = student_totals(student_questions).iloc[0]

        cols = st.columns(6)
        cols[0].metric("Classe", selected_total["classe"])
        cols[1].metric("Progetto", fmt_num(selected_total.get("progetto")))
        cols[2].metric("Totale esame", fmt_num(selected_total.get("totale_esame", selected_total["totale_integrato"])))
        cols[3].metric("Finale rounded", fmt_num(selected_total.get("finale_rounded")))
        cols[4].metric("Punti possibili", fmt_num(selected_total["punti_possibili"]))
        cols[5].metric("Da correggere", int(selected_total["domande_da_correggere"]))

        st.markdown("#### Scheda visione compiti")
        option_cols = st.columns([1, 1, 2])
        include_question_text = option_cols[0].checkbox("Includi testo domanda", value=False)
        include_correct_answer = option_cols[1].checkbox("Includi risposta corretta", value=False)
        review_docx = build_review_docx(
            selected_total,
            student_questions,
            include_question_text=include_question_text,
            include_correct_answer=include_correct_answer,
        )
        option_cols[2].download_button(
            "Scarica scheda visione compiti",
            data=review_docx,
            file_name=f"scheda_visione_compiti_{safe_filename(selected_total.get('full_name'))}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
        )

        review_questions = student_questions[student_questions["status"] != "Corretta"]
        st.markdown("#### Domande da rivedere")
        if review_questions.empty:
            st.success("Tutte le domande risultano corrette o gia valorizzate.")
        else:
            for _, row in review_questions.iterrows():
                with st.expander(
                    f"Domanda {int(row['question_num'])} - {row['status']} - {fmt_num(row['integrated_score'])}/{fmt_num(row['possible_points'])}",
                    expanded=row["status"] == "Da correggere",
                ):
                    st.markdown(status_html(row["status"]), unsafe_allow_html=True)
                    st.write("**Testo domanda**")
                    st.write(row["question_text"] or "-")
                    st.write("**Risposta studente**")
                    st.write(row["answer"] or "-")
                    if int(row["question_num"]) in {27, 28, 29} and row.get("comment"):
                        st.write("**Commento correzione**")
                        st.info(row["comment"])
                    st.caption(
                        f"Auto score: {fmt_num(row['auto_score'])} | Manual score salvato: {fmt_num(row['manual_score'])} | Punti: {fmt_num(row['possible_points'])}"
                    )

        st.markdown("#### Modifica manual score")
        edit_df = student_questions[
            [
                "question_num",
                "status",
                "question_id",
                "question_text",
                "answer",
                "comment",
                "possible_points",
                "auto_score",
                "manual_score",
                "integrated_score",
            ]
        ].copy()
        edited = st.data_editor(
            edit_df,
            key=f"student_editor_{selected_uid}",
            width="stretch",
            hide_index=True,
            disabled=[
                "question_num",
                "status",
                "question_id",
                "question_text",
                "answer",
                "comment",
                "possible_points",
                "auto_score",
                "integrated_score",
            ],
            column_config={
                "question_num": st.column_config.NumberColumn("Domanda", width="small"),
                "status": st.column_config.TextColumn("Stato", width="small"),
                "question_id": st.column_config.TextColumn("Question ID", width="small"),
                "question_text": st.column_config.TextColumn("Domanda", width="large"),
                "answer": st.column_config.TextColumn("Risposta", width="large"),
                "comment": st.column_config.TextColumn("Commento", width="medium"),
                "possible_points": st.column_config.NumberColumn("Punti", format="%.2f", width="small"),
                "auto_score": st.column_config.NumberColumn("Auto score", format="%.2f", width="small"),
                "manual_score": st.column_config.NumberColumn("Manual score", format="%.2f", min_value=0, width="small"),
                "integrated_score": st.column_config.NumberColumn("Punteggio integrato", format="%.2f", width="small"),
            },
        )

        errors = score_validation_errors(edited)
        if errors:
            for error in errors:
                st.error(error)
        if st.button("Salva correzioni studente", type="primary", disabled=bool(errors)):
            save_student_corrections(CORRECTIONS_PATH, student_questions, edited)
            st.success("Correzioni salvate.")
            st.cache_data.clear()
            st.rerun()

with tab_export:
    st.subheader("Export")
    export_totals = totals[
        [
            "classe",
            "sheet_name",
            "username",
            "last_name",
            "first_name",
            "full_name",
            "progetto",
            "totale_aperte",
            "totale_esame",
            "finale_non_arrotondato",
            "finale_rounded",
            "totale_autoscore",
            "totale_integrato",
            "punti_possibili",
            "domande_da_correggere",
            "domande_parziali",
            "domande_errate",
        ]
    ].rename(
        columns={
            "classe": "Classe",
            "sheet_name": "Foglio",
            "username": "Matricola",
            "last_name": "Cognome",
            "first_name": "Nome",
            "full_name": "Nome completo",
            "progetto": "Progetto",
            "totale_aperte": "Totale aperte",
            "totale_esame": "Totale esame",
            "finale_non_arrotondato": "Finale non arrotondato",
            "finale_rounded": "Finale rounded",
            "totale_autoscore": "Totale autoscore",
            "totale_integrato": "Totale integrato",
            "punti_possibili": "Punti possibili",
            "domande_da_correggere": "Domande da correggere",
            "domande_parziali": "Domande parziali",
            "domande_errate": "Domande errate",
        }
    )
    st.dataframe(export_totals, width="stretch", hide_index=True)
    st.download_button(
        "Scarica riepilogo voti CSV",
        data=export_totals.to_csv(index=False).encode("utf-8"),
        file_name="riepilogo_voti.csv",
        mime="text/csv",
    )

    corrected_bytes = build_corrected_workbook(
        source_for_export(path_text, uploaded_bytes),
        corrections,
        exam.exam_sheets,
    )
    st.download_button(
        "Scarica copia Excel con correzioni",
        data=corrected_bytes,
        file_name="esame_blackboard_corretto.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption(
        "La copia Excel compila i Manual Score salvati e aggiunge TOTALE AUTOSCORE APP, TOTALE INTEGRATO APP e DOMANDE DA CORREGGERE APP."
    )

    st.subheader("File pubblicazione voti ufficiale")
    official_uploaded = st.file_uploader(
        "Carica template Esiti ufficiale (.xls)",
        type=["xls"],
        key="official_results_upload",
    )
    official_path_text = st.text_input(
        "Oppure usa questo template locale",
        value=str(OFFICIAL_RESULTS_PATH),
    )
    official_source: bytes | Path | None = None
    if official_uploaded is not None:
        official_source = official_uploaded.getvalue()
        official_name = official_uploaded.name
    elif official_path_text:
        official_source = Path(official_path_text)
        official_name = Path(official_path_text).name
    else:
        official_name = "Esiti_compilato.xls"

    if official_source is not None:
        try:
            compiled_official, official_stats = compile_official_results_xls(official_source, totals)
            missing = official_stats["missing_in_template"]
            st.caption(
                f"Studenti compilati: {official_stats['matched']} su {official_stats['students_in_exam']} | "
                f"Colonne usate: Voto Q, Esame scritto W, Progetto X"
            )
            if missing:
                st.warning(
                    "Non trovati nel template ufficiale: "
                    + ", ".join(missing[:12])
                    + ("..." if len(missing) > 12 else "")
                )
            st.download_button(
                "Scarica Esiti ufficiale compilato (.xls)",
                data=compiled_official,
                file_name=f"{Path(official_name).stem}_compilato.xls",
                mime="application/vnd.ms-excel",
                type="primary",
            )
        except Exception as exc:
            st.error(f"Non riesco a compilare il file ufficiale: {exc}")
