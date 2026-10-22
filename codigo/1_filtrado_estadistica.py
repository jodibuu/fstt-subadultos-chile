"""Validación técnica y caracterización de series DICOM.

Selecciona una pila de TC geométricamente homogénea por carpeta, comprueba los
criterios de validez y registra las estadísticas de intensidad y su auditoría.
"""
from __future__ import annotations

import argparse
import os
import shutil
from pathlib import Path

import numpy as np
import pandas as pd
import SimpleITK as sitk
from sklearn.cluster import KMeans
from skimage.filters import threshold_otsu

from dicom_reader_robusto import DicomValidationError, cargar_volumen_dicom
from config_pipeline import DICOM_FOLDER, RESULTADOS_FOLDER


BASE_EXCEL = str(RESULTADOS_FOLDER / "resumen_final.xlsx")
OUTPUT_EXCEL = str(RESULTADOS_FOLDER / "resumen_final_v2.xlsx")
MIN_CORTES = 50
MAX_KMEANS = 200_000


COL_RESUMEN_BASE = [
    "Grupo", "Sexo", "Paciente", "Carpeta", "Slices", "Min", "Max", "Media", "Std",
    "P5", "P25", "P50", "P75", "P95", "Otsu",
    "KMeans_0", "KMeans_1", "KMeans_2",
]
AUDIT_COLS = [
    "Estado lectura", "Modalidad usada", "Matriz usada", "PhotometricInterpretation",
    "SamplesPerPixel", "Archivos detectados", "Archivos usados", "Archivos ignorados",
    "Grupos geométricos", "Pilas CT candidatas", "Duplicados de posición ignorados",
    "SeriesInstanceUID usado", "Advertencias",
]
COL_RESUMEN = COL_RESUMEN_BASE + AUDIT_COLS + ["KMeans_n_muestra"]
COL_VALIDOS = ["Grupo", "Sexo", "Paciente", "Carpeta", "Ruta"] + AUDIT_COLS
COL_ERRORES = ["Paciente", "Carpeta", "Codigo", "Motivo"] + AUDIT_COLS


def _norm(v) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def _empty_df(columns):
    return pd.DataFrame(columns=columns)


def cargar_hoja_segura(path: str, sheet: str, columns) -> pd.DataFrame:
    if not os.path.exists(path):
        return _empty_df(columns)
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception as exc:
        raise RuntimeError(
            f"No se pudo leer la hoja '{sheet}' del archivo Excel '{path}'. "
            "Se cancela el procesamiento para proteger los datos existentes."
        ) from exc
    for c in columns:
        if c not in df.columns:
            df[c] = np.nan
    return df


def ensure_output_from_base(base_excel: str, output_excel: str):
    os.makedirs(os.path.dirname(output_excel), exist_ok=True)
    if os.path.exists(output_excel):
        return
    if os.path.exists(base_excel):
        shutil.copy2(base_excel, output_excel)
        print(f"📋 Copia de trabajo creada desde: {base_excel}")
    else:
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            _empty_df(COL_RESUMEN).to_excel(writer, sheet_name="Resumen Estadístico", index=False)
            _empty_df(COL_ERRORES).to_excel(writer, sheet_name="Errores", index=False)
            _empty_df(COL_VALIDOS).to_excel(writer, sheet_name="Casos Válidos", index=False)
            _empty_df(["Paciente", "Carpeta"] + AUDIT_COLS).to_excel(writer, sheet_name="Auditoría DICOM", index=False)
        print(f"📋 Archivo nuevo creado: {output_excel}")


def muestra_determinista(values: np.ndarray, max_n: int = MAX_KMEANS) -> np.ndarray:
    values = np.asarray(values).reshape(-1)
    n = len(values)
    if n <= max_n:
        return values
    # Repartida a lo largo del vector; reproducible y sin RNG.
    idx = np.linspace(0, n - 1, max_n, dtype=np.int64)
    return values[idx]


def calcular_estadisticas(array: np.ndarray) -> dict:
    flat = np.asarray(array).reshape(-1)
    valid = flat[flat > -2000]
    if valid.size == 0:
        raise ValueError("Volumen vacío: no hay vóxeles > -2000 UH")

    p5, p25, p50, p75, p95 = np.percentile(valid, [5, 25, 50, 75, 95])
    muestra = muestra_determinista(valid)
    otsu_val = threshold_otsu(muestra)
    km = KMeans(n_clusters=3, random_state=0, n_init=10).fit(muestra.reshape(-1, 1))
    centers = sorted(float(x) for x in km.cluster_centers_.reshape(-1))

    return {
        "Min": float(np.min(valid)),
        "Max": float(np.max(valid)),
        "Media": float(np.mean(valid)),
        "Std": float(np.std(valid)),
        "P5": float(p5), "P25": float(p25), "P50": float(p50),
        "P75": float(p75), "P95": float(p95),
        "Otsu": float(otsu_val),
        "KMeans_0": centers[0], "KMeans_1": centers[1], "KMeans_2": centers[2],
        "KMeans_n_muestra": int(len(muestra)),
    }


def audit_subset(audit: dict) -> dict:
    return {c: audit.get(c, "") for c in AUDIT_COLS}


def localizar_series(pacientes: set[str] | None = None):
    """Yield (grupo, sexo, paciente, carpeta, dirpath) para carpetas con DICOM."""
    if not os.path.isdir(str(DICOM_FOLDER)):
        raise FileNotFoundError(f"No existe DICOM_FOLDER: {DICOM_FOLDER}")

    for grupo in sorted(os.listdir(str(DICOM_FOLDER))):
        gp = os.path.join(str(DICOM_FOLDER), grupo)
        if not os.path.isdir(gp):
            continue
        for sexo in sorted(os.listdir(gp)):
            sp = os.path.join(gp, sexo)
            if not os.path.isdir(sp):
                continue
        
            for paciente in sorted(os.listdir(sp)):
                pp = os.path.join(sp, paciente)
                if not os.path.isdir(pp):
                    continue
                if pacientes and paciente not in pacientes:
                    continue
                for dirpath, dirnames, filenames in os.walk(pp):
                    # Algunos DICOM carecen de extensión. Los directorios que
                    # contienen únicamente archivos comprimidos no constituyen
                    # por sí mismos una adquisición tomográfica.
                    archivos_candidatos = [
                        nombre for nombre in filenames
                        if not nombre.lower().endswith((".zip", ".7z", ".rar"))
                    ]
                    if archivos_candidatos:
                        yield grupo, sexo, paciente, os.path.basename(dirpath), dirpath


def quitar_par(df: pd.DataFrame, paciente: str, carpeta: str) -> pd.DataFrame:
    if df.empty or not {"Paciente", "Carpeta"}.issubset(df.columns):
        return df
    mask = ~(
        df["Paciente"].map(_norm).eq(paciente)
        & df["Carpeta"].map(_norm).eq(carpeta)
    )
    return df.loc[mask].copy()


def append_row(df: pd.DataFrame, row: dict) -> pd.DataFrame:
    return pd.concat([df, pd.DataFrame([row])], ignore_index=True, sort=False)


def procesar_una_serie(grupo, sexo, paciente, carpeta, dirpath):
    try:
        image, audit = cargar_volumen_dicom(dirpath, min_cortes=MIN_CORTES)
        array = sitk.GetArrayFromImage(image).astype(np.int16, copy=False)
        if array.ndim != 3:
            raise ValueError(f"Array no 3D después de lectura robusta: shape={array.shape}")

        stats = calcular_estadisticas(array)
        row_resumen = {
            "Grupo": grupo, "Sexo": sexo, "Paciente": paciente, "Carpeta": carpeta,
            "Slices": int(array.shape[0]), **stats, **audit_subset(audit),
        }
        row_valido = {
            "Grupo": grupo, "Sexo": sexo, "Paciente": paciente, "Carpeta": carpeta,
            "Ruta": dirpath, **audit_subset(audit),
        }
        return "ok", row_resumen, row_valido, audit

    except DicomValidationError as e:
        audit = e.audit or {}
        row_error = {
            "Paciente": paciente, "Carpeta": carpeta,
            "Codigo": e.code, "Motivo": str(e), **audit_subset(audit),
        }
        return "error", row_error, None, audit
    except Exception as e:
        audit = {"Estado lectura": "ERROR_NO_CLASIFICADO", "Advertencias": str(e)}
        row_error = {
            "Paciente": paciente, "Carpeta": carpeta,
            "Codigo": "ERROR_NO_CLASIFICADO", "Motivo": str(e), **audit_subset(audit),
        }
        return "error", row_error, None, audit


def construir_targets(modo: str, pacientes_args: list[str], errores_df: pd.DataFrame):
    pacientes = {p.strip() for p in pacientes_args if p.strip()}
    target_pairs = None

    if modo == "pacientes":
        if not pacientes:
            raw = input("Pacientes separados por coma (ej. CASO_001,CASO_002): ").strip()
            pacientes = {p.strip() for p in raw.split(",") if p.strip()}
        if not pacientes:
            raise SystemExit("No se indicó ningún paciente.")
    elif modo == "errores":
        if errores_df.empty:
            raise SystemExit("La hoja Errores está vacía.")
        target_pairs = set(zip(errores_df["Paciente"].map(_norm), errores_df["Carpeta"].map(_norm)))
    elif modo == "todos":
        pacientes = set()
    else:
        raise ValueError(modo)

    return pacientes or None, target_pairs


def main():
    parser = argparse.ArgumentParser(description="Filtrado DICOM robusto ")
    parser.add_argument("--modo", choices=["pacientes", "errores", "todos"], default=None)
    parser.add_argument("--pacientes", nargs="*", default=[])
    parser.add_argument("--base-excel", default=BASE_EXCEL)
    parser.add_argument("--salida", default=OUTPUT_EXCEL)
    args = parser.parse_args()

    modo = args.modo
    if modo is None:
        print("\n=== FILTRADO DICOM ROBUSTO  ===")
        print("1. Reprocesar pacientes concretos ")
        print("2. Reprocesar todas las filas actualmente en Errores")
        print("3. Reprocesar todas las carpetas DICOM")
        print("4. Salir")
        op = input("Opción: ").strip()
        modo = {"1": "pacientes", "2": "errores", "3": "todos"}.get(op)
        if op == "4":
            return
        if modo is None:
            raise SystemExit("Opción inválida")

    ensure_output_from_base(args.base_excel, args.salida)

    # Seguridad: antes de una corrida que pueda cambiar estados, crear respaldo.
    if os.path.exists(args.salida):
        from datetime import datetime
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = os.path.join(os.path.dirname(args.salida), f"resumen_final_v2_BACKUP_{ts}.xlsx")
        shutil.copy2(args.salida, backup)
        print(f"🛟 Respaldo automático: {backup}")

    resumen_df = cargar_hoja_segura(args.salida, "Resumen Estadístico", COL_RESUMEN)
    errores_df = cargar_hoja_segura(args.salida, "Errores", COL_ERRORES)
    casos_df = cargar_hoja_segura(args.salida, "Casos Válidos", COL_VALIDOS)
    audit_df = cargar_hoja_segura(args.salida, "Auditoría DICOM", ["Paciente", "Carpeta"] + AUDIT_COLS)

    # Limpia reintentos históricos antes de empezar.
    for df in (resumen_df, errores_df, casos_df, audit_df):
        if {"Paciente", "Carpeta"}.issubset(df.columns):
            df.drop_duplicates(subset=["Paciente", "Carpeta"], keep="last", inplace=True)

    pacientes, target_pairs = construir_targets(modo, args.pacientes, errores_df)
    if modo == "todos":
        print("⚠️  Modo TODOS recalcula estadísticas/K-means de todas las carpetas y puede tardar mucho.")

    targets = []
    for grupo, sexo, paciente, carpeta, dirpath in localizar_series(pacientes):
        if target_pairs is not None and (paciente, carpeta) not in target_pairs:
            continue
        targets.append((grupo, sexo, paciente, carpeta, dirpath))

    print(f"\nCarpetas a revisar: {len(targets)}")
    ok = err = recuperadas = 0

    for n, (grupo, sexo, paciente, carpeta, dirpath) in enumerate(targets, 1):
        print(f"[{n}/{len(targets)}] {paciente}/{carpeta}")
        status, row1, row2, audit = procesar_una_serie(grupo, sexo, paciente, carpeta, dirpath)

        # Reemplaza el estado histórico de esa pareja en las 4 tablas.
        resumen_df = quitar_par(resumen_df, paciente, carpeta)
        errores_df = quitar_par(errores_df, paciente, carpeta)
        casos_df = quitar_par(casos_df, paciente, carpeta)
        audit_df = quitar_par(audit_df, paciente, carpeta)

        audit_row = {"Paciente": paciente, "Carpeta": carpeta, **audit_subset(audit)}
        audit_df = append_row(audit_df, audit_row)

        if status == "ok":
            resumen_df = append_row(resumen_df, row1)
            casos_df = append_row(casos_df, row2)
            ok += 1
            if audit.get("Estado lectura") == "RECUPERADA_GEOMETRIA_MIXTA":
                recuperadas += 1
                print(f"  ✅ RECUPERADA: {audit.get('Matriz usada')} | usados={audit.get('Archivos usados')} | ignorados={audit.get('Archivos ignorados')}")
            else:
                print(f"  ✅ OK: {row1['Slices']} cortes")
        else:
            errores_df = append_row(errores_df, row1)
            err += 1
            print(f"  ❌ {row1['Codigo']}: {row1['Motivo']}")

    # Conflicto imposible: si quedó válido, no debe seguir como error.
    valid_pairs = set(zip(casos_df["Paciente"].map(_norm), casos_df["Carpeta"].map(_norm)))
    if not errores_df.empty:
        errores_df = errores_df.loc[
            [(p, c) not in valid_pairs for p, c in zip(errores_df["Paciente"].map(_norm), errores_df["Carpeta"].map(_norm))]
        ].copy()

    resumen_df.drop_duplicates(["Paciente", "Carpeta"], keep="last", inplace=True)
    errores_df.drop_duplicates(["Paciente", "Carpeta"], keep="last", inplace=True)
    casos_df.drop_duplicates(["Paciente", "Carpeta"], keep="last", inplace=True)
    audit_df.drop_duplicates(["Paciente", "Carpeta"], keep="last", inplace=True)

    v_pairs = set(zip(casos_df["Paciente"].map(_norm), casos_df["Carpeta"].map(_norm)))
    e_pairs = set(zip(errores_df["Paciente"].map(_norm), errores_df["Carpeta"].map(_norm)))
    pacientes_validos = {p for p, _ in v_pairs}
    resumen_global = pd.DataFrame([{
        "Total carpetas únicas": len(v_pairs | e_pairs),
        "Válidos únicos": len(v_pairs),
        "Errores/exclusiones únicas": len(e_pairs),
        "Pacientes con al menos una serie válida": len(pacientes_validos),
        "Recuperadas en esta corrida": recuperadas,
    }])

    with pd.ExcelWriter(args.salida, engine="openpyxl", mode="w") as writer:
        resumen_df.to_excel(writer, sheet_name="Resumen Estadístico", index=False)
        errores_df.to_excel(writer, sheet_name="Errores", index=False)
        casos_df.to_excel(writer, sheet_name="Casos Válidos", index=False)
        audit_df.to_excel(writer, sheet_name="Auditoría DICOM", index=False)
        resumen_global.to_excel(writer, sheet_name="Resumen Global", index=False)

    print("\n===== RESULTADO DE ESTA CORRIDA =====")
    print(f"OK: {ok}")
    print(f"Recuperadas por geometría mixta: {recuperadas}")
    print(f"Errores/exclusiones: {err}")
    print(f"Archivo actualizado: {args.salida}")


if __name__ == "__main__":
    main()
