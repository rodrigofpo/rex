"""Testes das fronteiras de segurança da interface web local."""

import json
import io
import os
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from email.message import Message

import pandas as pd
import pytest

from rex.errors import InvalidDocumentError
from rex.models import ExportPaths
from rex.ui.web import (
    HTML_PAGE,
    ExtractorWebHandler,
    dataframe_records_for_json,
    validate_upload_filename,
)


def test_web_page_uses_only_file_picker_for_report_selection():
    assert 'id="file-input"' in HTML_PAGE
    assert 'type="file"' in HTML_PAGE
    assert 'id="xlsx-options"' in HTML_PAGE
    assert re.search(r'<fieldset\s+id="xlsx-options"\s+disabled', HTML_PAGE)
    assert 'id="process-button"' in HTML_PAGE
    assert 'value="consolidated"' in HTML_PAGE
    assert 'value="per-sample"' in HTML_PAGE
    assert "/api/files" not in HTML_PAGE
    assert 'fetch("/api/extract-upload"' in HTML_PAGE
    assert '"X-REX-Workbook-Mode": state.mode' in HTML_PAGE
    assert 'href="/api/download/excel"' in HTML_PAGE
    assert 'href="/api/download/csv"' in HTML_PAGE
    assert "DADOS FICTICIOS" not in HTML_PAGE
    assert "setTimeout(finish" not in HTML_PAGE
    assert "new Blob(" not in HTML_PAGE


@pytest.mark.parametrize(("method", "path"), [("do_GET", "/api/files"), ("do_POST", "/api/extract")])
def test_web_disables_project_folder_selection_routes(method, path):
    handler = object.__new__(ExtractorWebHandler)
    handler.path = path
    handler.headers = Message()
    handler.headers["Host"] = "localhost"
    errors = []
    handler.send_error = lambda status, message: errors.append((status, message))

    getattr(handler, method)()

    assert errors[0][0] == 404


@pytest.mark.parametrize(
    ("failure", "expected_status", "expected_message"),
    [
        (InvalidDocumentError("DOCX inválido"), 422, "DOCX inválido"),
        (RuntimeError("detalhe interno"), 500, "Falha inesperada"),
    ],
)
def test_web_extraction_reports_controlled_errors(
    tmp_path: Path, monkeypatch, failure, expected_status, expected_message
):
    body = b"test"
    handler = object.__new__(ExtractorWebHandler)
    handler.target_dir = tmp_path
    handler.path = "/api/extract-upload"
    handler.rfile = io.BytesIO(body)
    handler.headers = Message()
    handler.headers["Host"] = "localhost:8085"
    handler.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    handler.headers["Content-Length"] = str(len(body))
    handler.headers["X-REX-Filename"] = "laudo.docx"
    responses = []
    handler._send_json = lambda status, payload: responses.append((status, payload))

    def fail_processing(*args, **kwargs):
        raise failure

    monkeypatch.setattr("rex.core.engine.EDSExtractorEngine.process", fail_processing)
    handler.do_POST()

    assert responses[0][0] == expected_status
    assert expected_message in responses[0][1]["error"]
    assert "detalhe interno" not in responses[0][1]["error"]


@pytest.mark.parametrize("filename", ["../outro.docx", "pasta/laudo.docx", "pasta\\laudo.docx", "laudo.txt", "~$laudo.docx", ""])
def test_upload_rejects_invalid_filename(filename):
    with pytest.raises(ValueError):
        validate_upload_filename(filename)


def test_upload_processes_selected_docx_and_removes_temporary_copy(tmp_path: Path, monkeypatch):
    content = b"conteudo-do-laudo"
    handler = object.__new__(ExtractorWebHandler)
    handler.target_dir = tmp_path
    handler.path = "/api/extract-upload"
    handler.rfile = io.BytesIO(content)
    handler.headers = Message()
    handler.headers["Host"] = "localhost:8085"
    handler.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    handler.headers["Content-Length"] = str(len(content))
    handler.headers["X-REX-Filename"] = "Laudo%20A.docx"
    handler.headers["X-REX-Workbook-Mode"] = "per-sample"
    responses = []
    handler._send_json = lambda status, payload: responses.append((status, payload))
    captured = {}

    def process(engine, docx_path, output_dir, workbook_mode):
        captured["name"] = docx_path.name
        captured["contents"] = docx_path.read_bytes()
        captured["path"] = docx_path
        captured["mode"] = workbook_mode
        captured["output_dir"] = output_dir
        engine.last_export_paths = ExportPaths(
            excel=tmp_path / "Laudo A_REX-teste.xlsx",
            csv=tmp_path / "Laudo A_REX-teste.csv",
        )
        return pd.DataFrame(), pd.DataFrame()

    monkeypatch.setattr("rex.core.engine.EDSExtractorEngine.process", process)
    handler.do_POST()

    assert responses[0][0] == 200
    assert captured["name"] == "Laudo A.docx"
    assert captured["contents"] == content
    assert captured["mode"] == "per-sample"
    assert captured["output_dir"] == tmp_path
    assert not captured["path"].exists()


def test_upload_rejects_oversized_body_before_reading(tmp_path: Path):
    handler = object.__new__(ExtractorWebHandler)
    handler.target_dir = tmp_path
    handler.path = "/api/extract-upload"
    handler.rfile = io.BytesIO(b"not read")
    handler.headers = Message()
    handler.headers["Host"] = "localhost"
    handler.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    handler.headers["Content-Length"] = str(250 * 1024 * 1024 + 1)
    responses = []
    handler._send_json = lambda status, payload: responses.append((status, payload))

    handler.do_POST()

    assert responses[0][0] == 400
    assert handler.rfile.tell() == 0


def test_upload_rejects_interrupted_body(tmp_path: Path):
    handler = object.__new__(ExtractorWebHandler)
    handler.target_dir = tmp_path
    handler.path = "/api/extract-upload"
    handler.rfile = io.BytesIO(b"short")
    handler.headers = Message()
    handler.headers["Host"] = "localhost"
    handler.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    handler.headers["Content-Length"] = "10"
    handler.headers["X-REX-Filename"] = "laudo.docx"
    responses = []
    handler._send_json = lambda status, payload: responses.append((status, payload))

    handler.do_POST()

    assert responses[0][0] == 400


class WebSerializationTests(unittest.TestCase):
    def test_converts_missing_dataframe_values_to_json_null(self):
        dataframe = pd.DataFrame(
            [{"PercentualPeso": 12.5, "SigmaPeso": float("nan"), "Data": pd.NaT}]
        )

        records = dataframe_records_for_json(dataframe)

        self.assertIsNone(records[0]["SigmaPeso"])
        self.assertIsNone(records[0]["Data"])
        self.assertEqual(json.loads(json.dumps(records, allow_nan=False)), records)

    def test_importing_web_module_does_not_load_ocr_runtime(self):
        project_root = Path(__file__).resolve().parents[1]
        environment = os.environ.copy()
        environment["PYTHONPATH"] = str(project_root)
        command = (
            "import sys; import rex.ui.web; "
            "assert 'rapidocr_onnxruntime' not in sys.modules; "
            "assert 'onnxruntime' not in sys.modules"
        )

        with tempfile.TemporaryDirectory() as directory:
            subprocess.run(
                [sys.executable, "-c", command],
                cwd=directory,
                env=environment,
                check=True,
            )


if __name__ == "__main__":
    unittest.main()
