"""Testes das fronteiras de segurança da interface web local."""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

from rex.ui.web import dataframe_records_for_json, resolve_docx_path


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
