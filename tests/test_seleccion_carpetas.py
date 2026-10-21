"""Verifica la selección de carpetas sin utilizar imágenes clínicas."""
import ast
import os
import tempfile
import unittest
from pathlib import Path


class SeleccionCarpetasTests(unittest.TestCase):
    def test_directorio_con_zip_no_equivale_a_adquisicion(self):
        source = Path(__file__).resolve().parents[1] / "codigo" / "1_filtrado_estadistica.py"
        module = ast.parse(source.read_text(encoding="utf-8"))
        func = next(n for n in module.body if isinstance(n, ast.FunctionDef) and n.name == "localizar_series")
        standalone = ast.Module(body=[func], type_ignores=[])
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            patient = root / "12-18" / "FEMENINO" / "CASO_001"
            patient.mkdir(parents=True)
            (patient / "copia.zip").write_bytes(b"archivo de ejemplo")
            (patient / "SERIE_A").mkdir()
            (patient / "SERIE_A" / "corte_001.dcm").write_bytes(b"ejemplo")
            (patient / "SERIE_B").mkdir()
            (patient / "SERIE_B" / "IMAGEN_SIN_EXTENSION").write_bytes(b"ejemplo")
            ns = {"os": os, "DICOM_FOLDER": root}
            exec(compile(standalone, str(source), "exec"), ns)
            encontrados = list(ns["localizar_series"]())
            self.assertEqual(sorted(item[3] for item in encontrados), ["SERIE_A", "SERIE_B"])


if __name__ == "__main__":
    unittest.main()
