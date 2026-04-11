import ast
import difflib
import html
import io
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

import pandas as pd

from ai_providers import AIProvider


def _load_similarity_namespace():
    source = Path("similarity_app.py").read_text(encoding="utf-8")
    module = ast.parse(source)
    needed_assigns = {
        "SIMILARITY_CONTEXT_PHRASES",
        "SIMILARITY_ABSTENTION_PATTERNS",
        "SIMILARITY_NEGATION_TOKENS",
        "SIMILARITY_TRUE_TOKENS",
        "SIMILARITY_FALSE_TOKENS",
        "SIMILARITY_STOPWORDS",
        "SIMILARITY_MONTH_TOKENS",
        "SIMILARITY_UNIT_SPECS",
        "SIMILARITY_UNIT_PATTERN",
        "SPREADSHEET_EXTENSIONS",
        "TEXT_FILE_EXTENSIONS",
        "JSON_FILE_EXTENSIONS",
        "PDF_FILE_EXTENSIONS",
        "WORD_FILE_EXTENSIONS",
        "LEGACY_WORD_FILE_EXTENSIONS",
        "COMPARE_ANY_TWO_EXTENSIONS",
    }
    needed_funcs = {
        "_normalize_similarity_text",
        "_strip_similarity_markup",
        "_prepare_answer_for_similarity",
        "_prepare_question_for_similarity",
        "_has_abstention_signal",
        "_char_ngrams",
        "_counter_cosine_similarity",
        "_lexical_similarity_percent_core",
        "_lexical_similarity_percent",
        "_has_negation_mismatch",
        "_extract_numeric_tokens",
        "_extract_percent_tokens",
        "_extract_date_markers",
        "_extract_boolean_label",
        "_extract_quantity_units",
        "_quantities_close",
        "_analyze_quantity_unit_relationships",
        "_significant_tokens",
        "_significant_overlap_ratio",
        "_is_short_subset_match",
        "_detect_similarity_conflicts",
        "_calibrate_similarity_score",
        "_uploaded_file_extension",
        "_is_spreadsheet_upload",
        "_read_csv_with_encodings",
        "_read_excel_upload",
        "_decode_uploaded_text",
        "_build_segment_dataframe",
        "_split_text_segments",
        "_stringify_nested_value",
        "_normalize_structured_dataframe",
        "_flatten_json_pairs",
        "_find_json_record_list",
        "_build_json_upload_dataframe",
        "_build_text_upload_dataframe",
        "_build_docx_upload_dataframe",
        "_extract_pdf_text_pages",
        "_build_pdf_upload_dataframe",
        "_default_compare_columns",
        "_should_use_best_match_alignment",
        "_should_use_question_context",
        "_alignment_seed_score",
        "_build_best_match_alignment",
        "read_uploaded_file",
        "_build_export_summary",
        "_build_non_excel_export_df",
        "_build_non_excel_json_payload",
        "_build_non_excel_html_payload",
    }

    chunks = []
    for node in module.body:
        if isinstance(node, ast.Assign):
            names = {target.id for target in node.targets if isinstance(target, ast.Name)}
            if names & needed_assigns:
                chunks.append(ast.get_source_segment(source, node))
        elif isinstance(node, ast.FunctionDef) and node.name in needed_funcs:
            chunks.append(ast.get_source_segment(source, node))

    namespace = {
        "re": re,
        "os": os,
        "json": json,
        "html": html,
        "io": io,
        "zipfile": zipfile,
        "ET": ET,
        "pd": pd,
        "difflib": difflib,
        "Counter": Counter,
    }
    exec("\n\n".join(chunks), namespace)
    return namespace


NS = _load_similarity_namespace()


class FakeUpload(io.BytesIO):
    def __init__(self, payload, name):
        super().__init__(payload)
        self.name = name


def _build_minimal_docx(paragraphs):
    xml_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        "<w:body>",
    ]
    for paragraph in paragraphs:
        xml_lines.append(f"<w:p><w:r><w:t>{paragraph}</w:t></w:r></w:p>")
    xml_lines.extend(["</w:body>", "</w:document>"])

    payload = io.BytesIO()
    with zipfile.ZipFile(payload, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", "".join(xml_lines))
    return payload.getvalue()


def test_extract_score_handles_common_formats():
    assert AIProvider.extract_score('{"score": 88}') == 88.0
    assert AIProvider.extract_score("score: 88 out of 100") == 88.0
    assert AIProvider.extract_score("The similarity is 85") == 85.0
    assert AIProvider.extract_score("0.92") == 92.0


def test_similarity_cleanup_and_abstention_detection_work():
    cleaned = NS["_prepare_answer_for_similarity"]("## Header\n- **Answer** [1][2]")
    assert "#" not in cleaned
    assert "**" not in cleaned
    assert "[1]" not in cleaned

    assert NS["_has_abstention_signal"]("The provided context does not specify this.") is True
    assert NS["_has_abstention_signal"]("The system supports ticket sales and reporting.") is False


def test_detect_similarity_conflicts_flags_real_answer_changes():
    conflicts = NS["_detect_similarity_conflicts"](
        "The application is approved after review.",
        "The application is not approved after review.",
    )
    assert conflicts["negation_mismatch"] is True

    conflicts = NS["_detect_similarity_conflicts"](
        "The order total is 25 items.",
        "The order total is 70 items.",
    )
    assert conflicts["numeric_mismatch"] is True

    conflicts = NS["_detect_similarity_conflicts"](
        "The due date is April 4, 2026.",
        "The due date is April 5, 2026.",
    )
    assert conflicts["date_mismatch"] is True

    conflicts = NS["_detect_similarity_conflicts"](
        "The package weighs 5 kg.",
        "The package weighs 5 lb.",
    )
    assert conflicts["unit_mismatch"] is True

    conflicts = NS["_detect_similarity_conflicts"](
        "The context does not provide that detail.",
        "The system uses a 22-inch touchscreen display.",
    )
    assert conflicts["abstention_mismatch"] is True


def test_detect_similarity_conflicts_allows_equivalent_unit_conversions():
    conflicts = NS["_detect_similarity_conflicts"](
        "The package weighs 5 kg.",
        "The package weighs 5000 g.",
    )
    assert conflicts["unit_mismatch"] is False
    assert conflicts["numeric_mismatch"] is False

    conflicts = NS["_detect_similarity_conflicts"](
        "The cable length is 2.54 cm.",
        "The cable length is 1 inch.",
    )
    assert conflicts["unit_mismatch"] is False


def test_balanced_calibration_caps_wrong_high_scores_and_boosts_matching_abstentions():
    negated = NS["_calibrate_similarity_score"](
        96,
        "The application is approved after review.",
        "The application is not approved after review.",
    )
    assert negated <= 62.0

    numeric = NS["_calibrate_similarity_score"](
        94,
        "The order total is 25 items.",
        "The order total is 70 items.",
    )
    assert numeric <= 68.0

    paraphrase = NS["_calibrate_similarity_score"](
        93,
        "The discount is 50%.",
        "The discount is 50 percent.",
    )
    assert paraphrase >= 90.0

    abstention = NS["_calibrate_similarity_score"](
        48,
        "The provided context does not state the main purpose of the terminal.",
        "The context does not specify the terminal''s main purpose.",
    )
    assert abstention >= 82.0

    partial_abstention = NS["_calibrate_similarity_score"](
        76,
        "The role is not explicitly detailed in the provided context, but it is listed in the system administration course agenda.",
        "The provided context does not describe the role in detail. It only indicates the role is included as a topic in the system administration course.",
    )
    assert partial_abstention >= 82.0


def test_non_excel_export_helpers_produce_expected_payloads():
    df = pd.DataFrame(
        [
            {"Question": "Q1", "Answer File1": "A", "Answer File2": "B", "Similarity": 87.5},
            {"Question": "Q2", "Answer File1": "C", "Answer File2": "D", "Similarity": 42.0},
        ]
    )

    summary = NS["_build_export_summary"](df, 85, "Similarity")
    assert summary["total_pairs"] == 2
    assert summary["above_threshold"] == 1
    assert summary["between_40_threshold"] == 1

    export_df = NS["_build_non_excel_export_df"](df)
    assert export_df.loc[0, "Similarity"] == "87.50%"

    json_payload = NS["_build_non_excel_json_payload"](df, summary)
    parsed = json.loads(json_payload)
    assert parsed["summary"]["average_similarity"] == 64.75
    assert len(parsed["rows"]) == 2

    html_payload = NS["_build_non_excel_html_payload"](export_df, summary, "Test Report")
    assert "<html" in html_payload.lower()
    assert "Test Report" in html_payload


def test_read_uploaded_file_supports_text_json_docx_and_excel():
    text_upload = FakeUpload(b"First line\nSecond line\n", "sample.txt")
    text_df = NS["read_uploaded_file"](text_upload)
    assert list(text_df.columns) == ["Section", "Content"]
    assert text_df.iloc[0]["Section"] == "Line 1"
    assert text_df.iloc[1]["Content"] == "Second line"

    json_upload = FakeUpload(
        json.dumps([{"question": "Q1", "answer": "A1"}]).encode("utf-8"),
        "sample.json",
    )
    json_df = NS["read_uploaded_file"](json_upload)
    assert list(json_df.columns) == ["question", "answer"]
    assert json_df.iloc[0]["answer"] == "A1"

    docx_upload = FakeUpload(_build_minimal_docx(["Alpha paragraph", "Beta paragraph"]), "sample.docx")
    docx_df = NS["read_uploaded_file"](docx_upload)
    assert list(docx_df.columns) == ["Section", "Content"]
    assert docx_df.iloc[0]["Section"] == "Paragraph 1"
    assert docx_df.iloc[1]["Content"] == "Beta paragraph"

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        pd.DataFrame([{"Question": "Q1", "Answer": "A1"}]).to_excel(writer, index=False)
    excel_upload = FakeUpload(excel_buffer.getvalue(), "sample.xlsx")
    excel_df = NS["read_uploaded_file"](excel_upload)
    assert list(excel_df.columns) == ["Question", "Answer"]
    assert excel_df.iloc[0]["Answer"] == "A1"


def test_default_compare_columns_uses_first_two_fields():
    doc_df = pd.DataFrame({"Section": ["Paragraph 1"], "Content": ["Body text"], "Extra": ["ignored"]})
    question_col, answer_col = NS["_default_compare_columns"](doc_df)
    assert question_col == "Section"
    assert answer_col == "Content"


def test_question_context_is_used_only_for_spreadsheet_flows():
    spreadsheet = FakeUpload(b"Question,Answer\nQ1,A1\n", "sample.csv")
    document = FakeUpload(b"Paragraph one", "sample.txt")

    assert NS["_should_use_question_context"]("Compare Two Columns in Same Excel File") is True
    assert NS["_should_use_question_context"]("Compare Two Excel Files") is True
    assert NS["_should_use_question_context"]("Compare Any Two Files", spreadsheet, spreadsheet) is True
    assert NS["_should_use_question_context"]("Compare Any Two Files", document, spreadsheet) is False


def test_best_match_alignment_reorders_document_sections_by_content():
    aligned_rows = NS["_build_best_match_alignment"](
        ["Page 1", "Page 2", "Page 3"],
        ["Alpha launch plan", "Beta shipping update", "Gamma return policy"],
        ["Page A", "Page B", "Page C"],
        ["Gamma return policy", "Alpha launch plan", "Beta shipping update"],
    )

    assert [row["question2"] for row in aligned_rows[:3]] == ["Page B", "Page C", "Page A"]
    assert [row["answer2"] for row in aligned_rows[:3]] == [
        "Alpha launch plan",
        "Beta shipping update",
        "Gamma return policy",
    ]


def test_build_pdf_upload_dataframe_reports_scanned_pdf_clearly():
    original = NS["_extract_pdf_text_pages"]
    try:
        NS["_extract_pdf_text_pages"] = lambda _file_bytes: ["", "   "]
        pdf_upload = FakeUpload(b"%PDF-1.4", "scan.pdf")
        try:
            NS["_build_pdf_upload_dataframe"](pdf_upload)
        except ValueError as exc:
            message = str(exc).lower()
            assert "scanned or image-only" in message
            assert "ocr" in message
        else:
            raise AssertionError("Expected scanned PDF uploads to raise a helpful ValueError")
    finally:
        NS["_extract_pdf_text_pages"] = original




