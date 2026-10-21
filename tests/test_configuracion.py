"""Prueba básica de la configuración sin imágenes clínicas."""
import importlib
import os
import sys
import tempfile
import unittest
from pathlib import Path

CODIGO = Path(__file__).resolve().parents[1] / "codigo"
sys.path.insert(0, str(CODIGO))


class ConfiguracionTests(unittest.TestCase):
    def test_directorios_y_recuperacion_ruta_historica(self):
        with tempfile.TemporaryDirectory() as tempdir:
            old = os.environ.get("FSTT_DATA_ROOT")
            os.environ["FSTT_DATA_ROOT"] = tempdir
            try:
                import config_pipeline
                cfg = importlib.reload(config_pipeline)
                folder = cfg.DICOM_FOLDER / "GRUPO" / "SEXO" / "CASO_001" / "SERIE_A"
                folder.mkdir(parents=True)
                row = {"Grupo":"GRUPO","Sexo":"SEXO","Paciente":"CASO_001",
                       "Carpeta":"SERIE_A","Ruta":"/unidad_anterior/1_DATOS_BRUTOS/GRUPO/SEXO/CASO_001/SERIE_A"}
                self.assertEqual(cfg.resolver_ruta_serie(row), folder)
            finally:
                if old is None:
                    os.environ.pop("FSTT_DATA_ROOT", None)
                else:
                    os.environ["FSTT_DATA_ROOT"] = old
                importlib.reload(cfg)


if __name__ == "__main__":
    unittest.main()
