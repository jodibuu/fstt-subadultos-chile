"""Lectura y validación de series DICOM geométricamente homogéneas."""
from __future__ import annotations

import math
import os
from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import SimpleITK as sitk


MIN_CORTES_DEFAULT = 50
SECONDARY_CAPTURE_PREFIX = "1.2.840.10008.5.1.4.1.1.7"


class DicomValidationError(RuntimeError):
    """Error DICOM con código estable y datos de auditoría."""

    def __init__(self, code: str, message: str, audit: Optional[dict] = None):
        super().__init__(message)
        self.code = code
        self.audit = audit or {}


@dataclass
class DicomRecord:
    path: str
    series_uid: str
    modality: str
    rows: int
    cols: int
    samples_per_pixel: int
    photometric: str
    sop_class_uid: str
    orientation: Tuple[float, ...]
    pixel_spacing: Tuple[float, ...]
    image_position: Tuple[float, ...]
    instance_number: Optional[int]
    number_of_frames: int
    convolution_kernel: str
    image_type: str

    @property
    def secondary_capture(self) -> bool:
        return self.sop_class_uid.startswith(SECONDARY_CAPTURE_PREFIX)

    @property
    def scalar_monochrome(self) -> bool:
        return (
            self.samples_per_pixel == 1
            and (self.photometric.startswith("MONOCHROME") or self.photometric == "")
        )

    @property
    def stack_key(self) -> tuple:
        """Características que deben ser homogéneas para apilar cortes."""
        return (
            self.series_uid,
            self.modality,
            self.rows,
            self.cols,
            self.samples_per_pixel,
            self.photometric,
            self.sop_class_uid,
            self.orientation,
            self.pixel_spacing,
            self.convolution_kernel,
        )


def _clean(value: Optional[str]) -> str:
    if value is None:
        return ""
    return str(value).replace("\x00", "").strip()


def _meta(reader: sitk.ImageFileReader, tag: str, default: str = "") -> str:
    try:
        if reader.HasMetaDataKey(tag):
            return _clean(reader.GetMetaData(tag))
    except Exception:
        pass
    return default


def _num_tuple(value: str, decimals: int = 5) -> Tuple[float, ...]:
    if not value:
        return tuple()
    parts = value.replace(",", ".").split("\\")
    out: List[float] = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        try:
            out.append(round(float(p), decimals))
        except ValueError:
            return tuple()
    return tuple(out)


def _int_or(value: str, default: int) -> int:
    try:
        return int(float(str(value).strip()))
    except Exception:
        return default


def leer_cabecera_dicom(path: str) -> DicomRecord:
    """Lee solamente la cabecera de un DICOM, sin cargar PixelData completo."""
    r = sitk.ImageFileReader()
    r.SetFileName(path)
    r.LoadPrivateTagsOff()
    r.ReadImageInformation()

    size = tuple(int(x) for x in r.GetSize())
    cols_default = size[0] if len(size) >= 1 else 0
    rows_default = size[1] if len(size) >= 2 else 0
    frames_default = size[2] if len(size) >= 3 else 1

    return DicomRecord(
        path=path,
        series_uid=_meta(r, "0020|000e"),
        modality=_meta(r, "0008|0060").upper(),
        rows=_int_or(_meta(r, "0028|0010"), rows_default),
        cols=_int_or(_meta(r, "0028|0011"), cols_default),
        samples_per_pixel=_int_or(_meta(r, "0028|0002"), 1),
        photometric=_meta(r, "0028|0004").upper(),
        sop_class_uid=_meta(r, "0008|0016"),
        orientation=_num_tuple(_meta(r, "0020|0037")),
        pixel_spacing=_num_tuple(_meta(r, "0028|0030")),
        image_position=_num_tuple(_meta(r, "0020|0032")),
        instance_number=(
            _int_or(_meta(r, "0020|0013"), 0)
            if _meta(r, "0020|0013") != ""
            else None
        ),
        number_of_frames=max(1, _int_or(_meta(r, "0028|0008"), frames_default)),
        convolution_kernel=_meta(r, "0018|1210").upper(),
        image_type=_meta(r, "0008|0008").upper(),
    )


def listar_dicom_directos(dirpath: str) -> List[str]:
    """Lista archivos DICOM de una carpeta; primero .dcm, con fallback prudente."""
    try:
        names = [n for n in os.listdir(dirpath) if os.path.isfile(os.path.join(dirpath, n))]
    except OSError as e:
        raise DicomValidationError("RUTA_NO_LEIBLE", f"No se pudo listar la carpeta: {e}")

    dcm = [os.path.join(dirpath, n) for n in names if n.lower().endswith(".dcm")]
    if dcm:
        return sorted(dcm)

    # Algunas exportaciones DICOM no usan extensión. Se prueban archivos normales;
    # los no-DICOM se descartarán al leer la cabecera.
    return sorted(os.path.join(dirpath, n) for n in names)


def _indices_muestra(n: int, max_muestras: int = 7) -> List[int]:
    """Índices deterministas repartidos entre inicio, centro y final."""
    if n <= 0:
        return []
    max_muestras = max(1, int(max_muestras))
    if n <= max_muestras:
        return list(range(n))
    if max_muestras == 1:
        return [0]
    # Linspace entero sin numpy: reparte muestras de 0 a n-1.
    idx = [round(i * (n - 1) / (max_muestras - 1)) for i in range(max_muestras)]
    return sorted(set(int(i) for i in idx))


def auditar_cabeceras_rapido(dirpath: str, max_muestras: int = 7) -> dict:
    """Auditoría rápida de metadatos SIN recorrer todos los cortes.

    Está pensada para revisar series ya clasificadas como válidas y detectar
    errores gruesos heredados (US/CR en vez de CT, RGB/multicomponente,
    Secondary Capture). NO sustituye a :func:`seleccionar_pila_ct` y nunca debe
    usarse para decidir por sí sola que una serie con geometría compleja es
    inválida.

    Solo lee cabeceras repartidas entre inicio, centro y final. Si alguna
    muestra no puede leerse, se informa, pero no se invalida la serie.
    """
    files = listar_dicom_directos(dirpath)
    audit = {
        "Ruta": dirpath,
        "Archivos detectados": len(files),
        "Cabeceras muestreadas": 0,
        "Cabeceras legibles": 0,
        "Cabeceras no legibles": 0,
        "Estado lectura": "",
        "Advertencias": "",
    }
    if not files:
        audit["Estado lectura"] = "SIN_ARCHIVOS"
        return audit

    sample = [files[i] for i in _indices_muestra(len(files), max_muestras=max_muestras)]
    records: List[DicomRecord] = []
    failures: List[str] = []
    for fp in sample:
        try:
            records.append(leer_cabecera_dicom(fp))
        except Exception as e:
            failures.append(f"{os.path.basename(fp)}: {e}")

    audit["Cabeceras muestreadas"] = len(sample)
    audit["Cabeceras legibles"] = len(records)
    audit["Cabeceras no legibles"] = len(failures)

    if not records:
        audit["Estado lectura"] = "MUESTRA_CABECERAS_NO_LEGIBLE"
        audit["Advertencias"] = " | ".join(failures[:3])
        return audit

    modalities = sorted({r.modality or "(vacío)" for r in records})
    audit["Modalidades detectadas"] = ", ".join(modalities)
    audit["Matrices muestreadas"] = ", ".join(sorted({f"{r.cols}x{r.rows}" for r in records}))
    audit["PhotometricInterpretation"] = ", ".join(sorted({r.photometric or "(vacío)" for r in records}))
    audit["SamplesPerPixel"] = ", ".join(str(x) for x in sorted({r.samples_per_pixel for r in records}))
    audit["SOPClassUID"] = ", ".join(sorted({r.sop_class_uid or "(vacío)" for r in records}))

    ct = [r for r in records if r.modality == "CT"]
    if not ct:
        audit["Estado lectura"] = "MODALIDAD_NO_CT_MUESTRA"
    else:
        segmentable = [r for r in ct if r.scalar_monochrome and not r.secondary_capture]
        if segmentable:
            # Si hay mezcla en la muestra, no invalida: solo pide revisión completa.
            mixed = len(segmentable) != len(ct) or len(ct) != len(records)
            audit["Estado lectura"] = "CT_APARENTA_SEGMENTABLE_MIXTA" if mixed else "CT_APARENTA_SEGMENTABLE"
        else:
            rgb = any(r.samples_per_pixel > 1 or r.photometric == "RGB" for r in ct)
            secondary = any(r.secondary_capture for r in ct)
            if rgb or secondary:
                audit["Estado lectura"] = "CT_NO_HU_MUESTRA"
            else:
                audit["Estado lectura"] = "CT_REVISAR_MUESTRA"

    warnings = []
    if failures:
        warnings.append(f"{len(failures)} cabecera(s) de muestra no legibles")
    if len({(r.rows, r.cols) for r in records}) > 1:
        warnings.append("matrices distintas en la muestra")
    if len({r.series_uid for r in records if r.series_uid}) > 1:
        warnings.append("más de un SeriesInstanceUID en la muestra")
    audit["Advertencias"] = "; ".join(warnings)
    return audit


def escanear_carpeta(dirpath: str) -> Tuple[List[DicomRecord], List[Tuple[str, str]]]:
    records: List[DicomRecord] = []
    fallos: List[Tuple[str, str]] = []
    for fp in listar_dicom_directos(dirpath):
        try:
            records.append(leer_cabecera_dicom(fp))
        except Exception as e:
            fallos.append((fp, str(e)))
    return records, fallos


def _cross(a: Sequence[float], b: Sequence[float]) -> Tuple[float, float, float]:
    return (
        a[1] * b[2] - a[2] * b[1],
        a[2] * b[0] - a[0] * b[2],
        a[0] * b[1] - a[1] * b[0],
    )


def _dot(a: Sequence[float], b: Sequence[float]) -> float:
    return float(sum(x * y for x, y in zip(a, b)))


def _slice_coordinate(rec: DicomRecord) -> Optional[float]:
    if len(rec.orientation) != 6 or len(rec.image_position) != 3:
        return None
    row = rec.orientation[:3]
    col = rec.orientation[3:]
    normal = _cross(row, col)
    norm = math.sqrt(_dot(normal, normal))
    if norm == 0:
        return None
    normal = tuple(v / norm for v in normal)
    return _dot(rec.image_position, normal)


def ordenar_y_deduplicar(records: Sequence[DicomRecord]) -> Tuple[List[DicomRecord], int, str]:
    """Ordena por posición anatómica y elimina posiciones duplicadas."""
    if not records:
        return [], 0, "sin_archivos"

    coords = [_slice_coordinate(r) for r in records]
    if all(c is not None for c in coords):
        pairs = sorted(zip(coords, records), key=lambda x: float(x[0]))
        selected: List[DicomRecord] = []
        last_coord: Optional[float] = None
        duplicates = 0
        for coord, rec in pairs:
            c = float(coord)
            if last_coord is not None and abs(c - last_coord) < 1e-3:
                duplicates += 1
                continue
            selected.append(rec)
            last_coord = c
        return selected, duplicates, "ImagePositionPatient"

    # Fallback por InstanceNumber; si falta, nombre de archivo.
    ordered = sorted(
        records,
        key=lambda r: (
            r.instance_number is None,
            r.instance_number if r.instance_number is not None else 10**12,
            os.path.basename(r.path),
        ),
    )
    return list(ordered), 0, "InstanceNumber/nombre"


def _candidate_summary(group: Sequence[DicomRecord]) -> dict:
    ordered, dup, order_method = ordenar_y_deduplicar(group)
    first = ordered[0] if ordered else group[0]
    multiframe = any(r.number_of_frames > 1 for r in group)
    if multiframe:
        effective = sum(max(1, r.number_of_frames) for r in group)
    else:
        effective = len(ordered)
    return {
        "records": ordered,
        "n_files": len(group),
        "n_effective_slices": int(effective),
        "duplicates_removed": int(dup),
        "order_method": order_method,
        "series_uid": first.series_uid,
        "modality": first.modality,
        "rows": first.rows,
        "cols": first.cols,
        "samples_per_pixel": first.samples_per_pixel,
        "photometric": first.photometric,
        "sop_class_uid": first.sop_class_uid,
        "secondary_capture": first.secondary_capture,
        "orientation": first.orientation,
        "pixel_spacing": first.pixel_spacing,
        "kernel": first.convolution_kernel,
        "image_type": first.image_type,
        "multiframe": multiframe,
    }


def detectar_pilas(records: Sequence[DicomRecord]) -> List[dict]:
    groups: Dict[tuple, List[DicomRecord]] = defaultdict(list)
    for rec in records:
        groups[rec.stack_key].append(rec)
    return [_candidate_summary(group) for group in groups.values()]


def _is_segmentable_candidate(c: dict) -> bool:
    photo = str(c["photometric"]).upper()
    return (
        str(c["modality"]).upper() == "CT"
        and int(c["samples_per_pixel"]) == 1
        and (photo.startswith("MONOCHROME") or photo == "")
        and not bool(c["secondary_capture"])
    )


def seleccionar_pila_ct(
    dirpath: str,
    min_cortes: int = MIN_CORTES_DEFAULT,
) -> Tuple[dict, dict]:
    """Selecciona la mejor pila CT escalar homogénea de una carpeta.

    Devuelve ``(candidato, auditoria)``. Si no existe una pila compatible,
    lanza ``DicomValidationError`` con una razón estable.
    """
    records, header_failures = escanear_carpeta(dirpath)
    audit = {
        "Ruta": dirpath,
        "Archivos detectados": len(records) + len(header_failures),
        "Archivos con cabecera DICOM": len(records),
        "Cabeceras no legibles": len(header_failures),
        "Grupos geométricos": 0,
        "Estado lectura": "",
        "Advertencias": "",
    }

    if not records:
        audit["Estado lectura"] = "SIN_DICOM_LEGIBLE"
        raise DicomValidationError(
            "SIN_DICOM_LEGIBLE",
            "No se encontró ningún DICOM legible en la carpeta.",
            audit,
        )

    modalities = sorted({r.modality or "(vacío)" for r in records})
    audit["Modalidades detectadas"] = ", ".join(modalities)

    ct_records = [r for r in records if r.modality == "CT"]
    if not ct_records:
        audit["Estado lectura"] = "MODALIDAD_NO_CT"
        raise DicomValidationError(
            "MODALIDAD_NO_CT",
            "La carpeta no contiene una serie de tomografía computarizada (CT). "
            f"Modalidades detectadas: {', '.join(modalities)}",
            audit,
        )

    candidates = detectar_pilas(ct_records)
    audit["Grupos geométricos"] = len(candidates)

    valid_geometry = [c for c in candidates if _is_segmentable_candidate(c)]
    if not valid_geometry:
        rgb = any(int(c["samples_per_pixel"]) > 1 or str(c["photometric"]).upper() == "RGB" for c in candidates)
        secondary = any(bool(c["secondary_capture"]) for c in candidates)
        audit["Estado lectura"] = "CT_NO_HU"
        details = []
        if rgb:
            details.append("RGB/multicomponente")
        if secondary:
            details.append("Secondary Capture")
        detail = ", ".join(details) if details else "no escalar/monocromática"
        raise DicomValidationError(
            "CT_NO_HU",
            "Hay archivos con modalidad CT, pero ninguna pila escalar compatible "
            f"con segmentación en HU ({detail}).",
            audit,
        )

    # Prioridad: más cortes; luego mayor matriz. Es determinista.
    valid_geometry.sort(
        key=lambda c: (c["n_effective_slices"], c["rows"] * c["cols"], c["n_files"]),
        reverse=True,
    )
    chosen = valid_geometry[0]

    audit.update({
        "SeriesInstanceUID usado": chosen["series_uid"],
        "Modalidad usada": chosen["modality"],
        "Matriz usada": f"{chosen['cols']}x{chosen['rows']}",
        "PhotometricInterpretation": chosen["photometric"],
        "SamplesPerPixel": chosen["samples_per_pixel"],
        "Archivos usados": chosen["n_files"],
        "Cortes efectivos": chosen["n_effective_slices"],
        "Archivos ignorados": len(records) - chosen["n_files"],
        "Duplicados de posición ignorados": chosen["duplicates_removed"],
        "Método de orden": chosen["order_method"],
        "Pilas CT candidatas": len(valid_geometry),
    })

    warnings: List[str] = []
    if len(candidates) > 1:
        warnings.append(
            f"Se detectaron {len(candidates)} geometrías; se eligió la pila homogénea más completa."
        )
    if len(valid_geometry) > 1:
        warnings.append(
            f"Había {len(valid_geometry)} pilas CT segmentables; revisar visualmente la elegida."
        )
    if audit["Archivos ignorados"]:
        warnings.append(f"Se ignoraron {audit['Archivos ignorados']} archivos incompatibles con la pila elegida.")
    if chosen["duplicates_removed"]:
        warnings.append(f"Se ignoraron {chosen['duplicates_removed']} posiciones duplicadas.")
    audit["Advertencias"] = " ".join(warnings)

    if chosen["n_effective_slices"] < int(min_cortes):
        audit["Estado lectura"] = "CT_MENOS_50_CORTES"
        raise DicomValidationError(
            "CT_MENOS_50_CORTES",
            f"La mejor pila CT homogénea tiene {chosen['n_effective_slices']} cortes; "
            f"se requieren al menos {min_cortes}.",
            audit,
        )

    audit["Estado lectura"] = (
        "RECUPERADA_GEOMETRIA_MIXTA" if len(candidates) > 1 or audit["Archivos ignorados"] else "OK"
    )
    return chosen, audit


def cargar_volumen_dicom(
    dirpath: str,
    min_cortes: int = MIN_CORTES_DEFAULT,
):
    """Carga una pila CT 3D homogénea y devuelve ``image, audit``.

    No redimensiona ni interpola. Para un DICOM multiframe único usa ReadImage;
    para una serie convencional usa ImageSeriesReader solo con los archivos de
    la pila seleccionada.
    """
    chosen, audit = seleccionar_pila_ct(dirpath, min_cortes=min_cortes)
    records: List[DicomRecord] = chosen["records"]

    try:
        if len(records) == 1 and records[0].number_of_frames > 1:
            image = sitk.ReadImage(records[0].path)
        else:
            reader = sitk.ImageSeriesReader()
            reader.SetFileNames([r.path for r in records])
            image = reader.Execute()
    except Exception as e:
        audit["Estado lectura"] = "FALLO_LECTURA_PILA_SELECCIONADA"
        raise DicomValidationError(
            "FALLO_LECTURA_PILA_SELECCIONADA",
            f"La pila homogénea fue identificada, pero SimpleITK no pudo cargarla: {e}",
            audit,
        ) from e

    dim = int(image.GetDimension())
    comps = int(image.GetNumberOfComponentsPerPixel())
    size = tuple(int(v) for v in image.GetSize())
    audit["Dimensión SimpleITK"] = dim
    audit["Componentes por pixel SimpleITK"] = comps
    audit["Size SimpleITK"] = "x".join(map(str, size))

    if dim != 3 or comps != 1:
        audit["Estado lectura"] = "VOLUMEN_NO_3D_ESCALAR"
        raise DicomValidationError(
            "VOLUMEN_NO_3D_ESCALAR",
            f"La pila seleccionada no produjo un volumen escalar 3D: dimension={dim}, componentes={comps}.",
            audit,
        )

    return image, audit
