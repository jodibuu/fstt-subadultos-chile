"""Configuración de rutas del proyecto y resolución de entradas DICOM."""
from __future__ import annotations

import os
from pathlib import Path


def _path_env(name: str, default: Path) -> Path:
    return Path(os.environ.get(name, str(default))).expanduser().resolve()


# Por defecto se utiliza una carpeta local, excluida mediante .gitignore.
BASE_PATH = _path_env("FSTT_DATA_ROOT", Path(__file__).resolve().parents[1] / "datos_locales")
DICOM_FOLDER = _path_env("FSTT_DICOM_DIR", BASE_PATH / "1_DATOS_BRUTOS")
RESULTADOS_FOLDER = _path_env("FSTT_RESULTADOS_DIR", BASE_PATH / "5_RESULTADOS")
SEGMENT_FOLDER = _path_env("FSTT_SEGMENT_DIR", BASE_PATH / "2_PROCESAMIENTO" / "SEGMENTACIONES")
MALLAS_FOLDER = _path_env("FSTT_MALLAS_DIR", BASE_PATH / "3_MALLAS_PLY")


def norm(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def _normalizar_path_texto(value: str) -> str:
    return value.strip().strip('"').replace("\\", "/")


def resolver_ruta_serie(row, devolver_intentos: bool = False):
    """Localiza una serie sin modificar los metadatos ni reclasificar el caso.

    Prueba la ruta registrada, una reconstrucción por sufijo de datos brutos,
    la estructura Grupo/Sexo/Paciente/Carpeta y una búsqueda local dentro de
    la carpeta del paciente. Devuelve ``Path`` o ``None``; opcionalmente
    también la lista de rutas examinadas.
    """
    def get(key):
        try:
            return norm(row[key])
        except (KeyError, TypeError, AttributeError):
            return norm(getattr(row, key, ""))

    ruta_excel = _normalizar_path_texto(get("Ruta")) if get("Ruta") else ""
    grupo, sexo = get("Grupo"), get("Sexo")
    paciente, carpeta = get("Paciente"), get("Carpeta")

    candidates = []
    seen = set()

    def add(value):
        if not value:
            return
        path_text = _normalizar_path_texto(str(value))
        if path_text and path_text.lower() not in seen:
            seen.add(path_text.lower())
            candidates.append(Path(path_text))

    add(ruta_excel)
    if ruta_excel:
        # Reconstruye rutas antiguas sin codificar letras de unidad ni usuarios.
        needle = f"/{DICOM_FOLDER.name.lower()}/"
        normalized = "/" + ruta_excel.lstrip("/")
        idx = normalized.lower().find(needle)
        if idx >= 0:
            suffix = normalized[idx + len(needle):]
            add(DICOM_FOLDER / Path(suffix))

    if grupo and sexo and paciente and carpeta:
        add(DICOM_FOLDER / grupo / sexo / paciente / carpeta)

    for path in candidates:
        if path.is_dir():
            return (path, [str(c) for c in candidates]) if devolver_intentos else path

    if grupo and sexo and paciente and carpeta:
        patient_dir = DICOM_FOLDER / grupo / sexo / paciente
        if patient_dir.is_dir():
            for current, _, _ in os.walk(patient_dir):
                if Path(current).name == carpeta:
                    found = Path(current)
                    attempts = [str(c) for c in candidates] + [str(found)]
                    return (found, attempts) if devolver_intentos else found

    return (None, [str(c) for c in candidates]) if devolver_intentos else None
