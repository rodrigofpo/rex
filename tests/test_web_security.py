"""Testes das fronteiras de segurança da interface web local."""

import json
import io
import os
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
    ExtractorWebHandler,
    dataframe_records_for_json,
    resolve_docx_path,
    validate_upload_filename,
)


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
    (tmp_path / "laudo.docx").write_bytes(b"test")
    body = json.dumps({"filename": "laudo.docx"}).encode("utf-8")
    handler = object.__new__(ExtractorWebHandler)
    handler.target_dir = tmp_path
    handler.path = "/api/extract"
    handler.rfile = io.BytesIO(body)
    handler.headers = Message()
    handler.headers["Host"] = "localhost:8085"
    handler.headers["Content-Type"] = "application/json"
    handler.headers["Content-Length"] = str(len(body))
    responses = []
    handler._send_json = lambda status, payload: responses.append((status, payload))

    def fail_processing(*args, **kwargs):
        raise failure

    monkeypatch.setattr("rex.core.engine.EDSExtractorEngine.process", fail_processing)
    handler.do_POST()

    assert responses[0][0] == expected_status
    assert expected_message in responses[0][1]["error"]
    assert "detalhe interno" not in responses[0][1]["error"]


class ResolveDocxPathTests(unittest.TestCase):
    def test_accepts_file_inside_target(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory)
            report = target / "report.docx"
            report.write_bytes(b"test")

            self.assertEqual(resolve_docx_path(target, report.name), report.resolve())

    def test_rejects_invalid_or_external_paths(self):
        invalid_names = (
            None,
            "",
            "../secret.docx",
            "/tmp/secret.docx",
            "report.txt",
            "folder/report.docx",
        )
        with tempfile.TemporaryDirectory() as directory:
            for filename in invalid_names:
                with self.subTest(filename=filename):
                    with self.assertRaises(ValueError):
                        resolve_docx_path(directory, filename)

    def test_rejects_symlink_escape(self):
        with tempfile.TemporaryDirectory() as parent_directory:
            parent = Path(parent_directory)
            target = parent / "published"
            target.mkdir()
            outside = parent / "outside.docx"
            outside.write_bytes(b"test")
            link = target / "linked.docx"
            link.symlink_to(outside)

            with self.assertRaises(ValueError):
                resolve_docx_path(target, link.name)


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
