"""Segmentación volumétrica de tejido óseo y blando.

Delimita la región corporal mediante un umbral fijo y emplea el percentil 95
de las intensidades válidas del volumen como umbral óseo individualizado.
"""
from __future__ import annotations

import argparse
import os
import shutil
from datetime import datetime

import numpy as np
import pandas as pd
import SimpleITK as sitk
from scipy.ndimage import binary_closing, binary_fill_holes
from sklearn.cluster import KMeans
from skimage import measure
from skimage.filters import threshold_otsu

from dicom_reader_robusto import DicomValidationError, cargar_volumen_dicom, auditar_cabeceras_rapido
from config_pipeline import RESULTADOS_FOLDER, SEGMENT_FOLDER, resolver_ruta_serie


CASOS_XLSX = str(RESULTADOS_FOLDER / "resumen_final_v2.xlsx")
BASE_SEG_XLSX = str(RESULTADOS_FOLDER / "estadisticas_segmentacion.xlsx")
SALIDA_XLSX = str(RESULTADOS_FOLDER / "estadisticas_segmentacion_v2.xlsx")
LOG_PATH = str(RESULTADOS_FOLDER / "log_segmentacion_v3.txt")
MIN_CORTES = 50
MAX_KMEANS = 200_000

os.makedirs(RESULTADOS_FOLDER, exist_ok=True)
os.makedirs(SEGMENT_FOLDER, exist_ok=True)

AUDIT_COLS = [
    "Estado lectura", "Modalidad usada", "Matriz usada", "PhotometricInterpretation",
    "SamplesPerPixel", "Archivos detectados", "Archivos usados", "Archivos ignorados",
    "Grupos geométricos", "Pilas CT candidatas", "Duplicados de posición ignorados",
    "SeriesInstanceUID usado", "Advertencias",
]

COLUMNAS = [
    "Grupo", "Sexo", "Paciente", "Carpeta", "Slices", "Min", "Max", "Media", "Std",
    "P5", "P25", "P50", "P75", "P95", "Otsu",
    "KMeans_0", "KMeans_1", "KMeans_2", "KMeans_n_muestra",
    "Volumen total (mm3)", "Volumen blando (mm3)", "Volumen hueso (mm3)",
    "% blando", "% hueso", "Ruta salida segmentación",
] + AUDIT_COLS
COLUMNAS_ERRORES = ["Paciente", "Carpeta", "Codigo", "Motivo"] + AUDIT_COLS


def log_evento(msg: str):
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now().isoformat()} - {msg}\n")


def _norm(v) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def resolver_ruta_dicom(row):
    ruta = resolver_ruta_serie(row)
    return str(ruta) if ruta is not None else ""


def estandarizar_clave(valor) -> str:
    return str(valor).strip().replace(".", "_")


def ensure_output(base_path: str, out_path: str):
    if os.path.exists(out_path):
        return
    if os.path.exists(base_path):
        shutil.copy2(base_path, out_path)
        print(f"📋 Copia de trabajo creada desde {base_path}")
    else:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            pd.DataFrame(columns=COLUMNAS).to_excel(writer, sheet_name="Estadísticas", index=False)
            pd.DataFrame(columns=COLUMNAS_ERRORES).to_excel(writer, sheet_name="Errores", index=False)


def cargar_hoja(path, sheet, columns):
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


def quitar_par(df, paciente, carpeta):
    if df.empty or not {"Paciente", "Carpeta"}.issubset(df.columns):
        return df
    mask = ~(
        df["Paciente"].map(_norm).eq(paciente)
        & df["Carpeta"].map(_norm).eq(carpeta)
    )
    return df.loc[mask].copy()


def muestra_determinista(values: np.ndarray, max_n: int = MAX_KMEANS) -> np.ndarray:
    values = np.asarray(values).reshape(-1)
    if len(values) <= max_n:
        return values
    idx = np.linspace(0, len(values) - 1, max_n, dtype=np.int64)
    return values[idx]


def estadisticas(array: np.ndarray) -> dict:
    flat = array.reshape(-1)
    valid = flat[flat > -2000]
    if valid.size == 0:
        raise ValueError("Volumen vacío")
    p5, p25, p50, p75, p95 = np.percentile(valid, [5, 25, 50, 75, 95])
    muestra = muestra_determinista(valid)
    otsu = threshold_otsu(muestra)
    km = KMeans(n_clusters=3, random_state=0, n_init=10).fit(muestra.reshape(-1, 1))
    centers = sorted(float(x) for x in km.cluster_centers_.reshape(-1))
    return {
        "Min": float(np.min(valid)), "Max": float(np.max(valid)),
        "Media": float(np.mean(valid)), "Std": float(np.std(valid)),
        "P5": float(p5), "P25": float(p25), "P50": float(p50),
        "P75": float(p75), "P95": float(p95), "Otsu": float(otsu),
        "KMeans_0": centers[0], "KMeans_1": centers[1], "KMeans_2": centers[2],
        "KMeans_n_muestra": int(len(muestra)),
    }


def limpiar_mascara(mask: np.ndarray) -> np.ndarray:
    labels = measure.label(mask)
    props = measure.regionprops(labels)
    if not props:
        return np.zeros_like(mask, dtype=np.uint8)
    mayor = max(props, key=lambda x: x.area).label
    return (labels == mayor).astype(np.uint8)


def eliminar_pequenas_componentes(mask: np.ndarray, volumen_min: int) -> np.ndarray:
    labels = measure.label(mask)
    props = measure.regionprops(labels)
    result = np.zeros_like(mask, dtype=np.uint8)
    for p in props:
        if p.area >= volumen_min:
            result[labels == p.label] = 1
    return result


def generar_mascara_cuerpo(array: np.ndarray, umbral_aire: int = -500) -> np.ndarray:
    cuerpo = array >= umbral_aire
    cuerpo = binary_fill_holes(cuerpo)
    cuerpo = limpiar_mascara(cuerpo.astype(np.uint8))
    cuerpo = binary_closing(cuerpo, structure=np.ones((5, 5, 5), dtype=bool))
    return cuerpo.astype(np.uint8)


def guardar_nifti(array: np.ndarray, ref_img: sitk.Image, path: str):
    img_out = sitk.GetImageFromArray(array)
    img_out.CopyInformation(ref_img)
    sitk.WriteImage(img_out, path)


def crear_bat_itksnap(out_dir: str, nombre_base: str):
    bat_filename = f"abrir_{nombre_base.split('_')[0]}_itksnap.bat"
    bat_path = os.path.join(out_dir, bat_filename)
    path_g = os.path.abspath(os.path.join(out_dir, f"{nombre_base}_original.nii.gz")).replace("\\", "/")
    path_s = os.path.abspath(os.path.join(out_dir, f"{nombre_base}_segmentacion.label.nii.gz")).replace("\\", "/")
    with open(bat_path, "w", encoding="utf-8") as bat:
        bat.write("@echo off\n")
        bat.write("REM Abrir volumen + segmentación en ITK-SNAP\n")
        bat.write('set ITKSNAP="%ITKSNAP_EXECUTABLE%"\n')
        bat.write(f'%ITKSNAP% -g "{path_g}" -s "{path_s}"\n')
        bat.write("pause\n")


def audit_subset(audit: dict) -> dict:
    return {c: audit.get(c, "") for c in AUDIT_COLS}


def segmentar_caso(row) -> tuple[str, dict, dict | None]:
    grupo = _norm(row["Grupo"])
    sexo = _norm(row["Sexo"])
    paciente = estandarizar_clave(row["Paciente"])
    carpeta = estandarizar_clave(row["Carpeta"])
    nombre_base = f"{paciente}_{carpeta}"
    out_dir = os.path.join(str(SEGMENT_FOLDER), grupo, sexo, paciente, nombre_base)
    os.makedirs(out_dir, exist_ok=True)

    try:
        ruta_dicom = resolver_ruta_dicom(row)
        if not os.path.isdir(ruta_dicom):
            raise DicomValidationError("RUTA_NO_EXISTE", f"Ruta DICOM no existe: {ruta_dicom}")

        image, audit = cargar_volumen_dicom(ruta_dicom, min_cortes=MIN_CORTES)
        array = sitk.GetArrayFromImage(image).astype(np.int16, copy=False)
        if array.ndim != 3:
            raise ValueError(f"Array no 3D: shape={array.shape}")

        st = estadisticas(array)
        p95 = st["P95"]
        mask = generar_mascara_cuerpo(array, -500)

        guardar_nifti(array * mask, image, os.path.join(out_dir, f"{nombre_base}_original.nii.gz"))
        guardar_nifti(mask, image, os.path.join(out_dir, f"{nombre_base}_cuerpo.nii.gz"))

        aire = ((mask == 0) | (array < -500)).astype(np.uint8)
        hueso = ((mask == 1) & (array >= p95)).astype(np.uint8)
        hueso = eliminar_pequenas_componentes(limpiar_mascara(hueso), 500)
        blando = ((mask == 1) & (aire == 0) & (hueso == 0)).astype(np.uint8)
        blando = eliminar_pequenas_componentes(limpiar_mascara(blando), 2000)

        label = np.zeros_like(array, dtype=np.uint8)
        label[blando == 1] = 1
        label[hueso == 1] = 2

        seg_path = os.path.join(out_dir, f"{nombre_base}_segmentacion.label.nii.gz")
        guardar_nifti(label, image, seg_path)
        crear_bat_itksnap(out_dir, nombre_base)

        spacing = image.GetSpacing()  # x,y,z
        voxel_vol = float(spacing[0] * spacing[1] * spacing[2])
        vol_total = float(np.sum(mask) * voxel_vol)
        vol_blando = float(np.sum(blando) * voxel_vol)
        vol_hueso = float(np.sum(hueso) * voxel_vol)

        out = {
            "Grupo": grupo, "Sexo": sexo, "Paciente": paciente, "Carpeta": carpeta,
            "Slices": int(array.shape[0]), **st,
            "Volumen total (mm3)": round(vol_total, 2),
            "Volumen blando (mm3)": round(vol_blando, 2),
            "Volumen hueso (mm3)": round(vol_hueso, 2),
            "% blando": round(100 * vol_blando / vol_total, 2) if vol_total else 0,
            "% hueso": round(100 * vol_hueso / vol_total, 2) if vol_total else 0,
            "Ruta salida segmentación": seg_path,
            **audit_subset(audit),
        }
        return "ok", out, audit

    except DicomValidationError as e:
        audit = e.audit or {}
        return "error", {
            "Paciente": paciente, "Carpeta": carpeta,
            "Codigo": e.code, "Motivo": str(e), **audit_subset(audit),
        }, audit
    except Exception as e:
        audit = {"Estado lectura": "ERROR_SEGMENTACION", "Advertencias": str(e)}
        return "error", {
            "Paciente": paciente, "Carpeta": carpeta,
            "Codigo": "ERROR_SEGMENTACION", "Motivo": str(e), **audit_subset(audit),
        }, audit


def seg_path_for(row) -> str:
    grupo = _norm(row["Grupo"])
    sexo = _norm(row["Sexo"])
    paciente = estandarizar_clave(row["Paciente"])
    carpeta = estandarizar_clave(row["Carpeta"])
    nombre_base = f"{paciente}_{carpeta}"
    return os.path.join(
        str(SEGMENT_FOLDER), grupo, sexo, paciente, nombre_base,
        f"{nombre_base}_segmentacion.label.nii.gz",
    )


def clave_par(paciente, carpeta):
    return (_norm(paciente), estandarizar_clave(carpeta))


def pares_df(df):
    if df.empty or not {"Paciente", "Carpeta"}.issubset(df.columns):
        return set()
    return {clave_par(p, c) for p, c in zip(df["Paciente"], df["Carpeta"])}


def guardar_estado_excel(path, df_estad, df_err, ok_corrida, err_corrida):
    df_estad = df_estad.copy()
    df_err = df_err.copy()
    df_estad.drop_duplicates(["Paciente", "Carpeta"], keep="last", inplace=True)
    df_err.drop_duplicates(["Paciente", "Carpeta"], keep="last", inplace=True)

    ok_pairs = pares_df(df_estad)
    if not df_err.empty:
        keep = [clave_par(p, c) not in ok_pairs for p, c in zip(df_err["Paciente"], df_err["Carpeta"])]
        df_err = df_err.loc[keep].copy()
    err_pairs = pares_df(df_err)

    resumen = pd.DataFrame([{
        "Segmentaciones registradas": len(ok_pairs),
        "Errores vigentes": len(err_pairs),
        "Total pares registrados": len(ok_pairs | err_pairs),
        "Procesadas OK en esta corrida": ok_corrida,
        "Errores en esta corrida": err_corrida,
    }])

    tmp = path + ".tmp.xlsx"
    with pd.ExcelWriter(tmp, engine="openpyxl", mode="w") as writer:
        df_estad.to_excel(writer, sheet_name="Estadísticas", index=False)
        df_err.to_excel(writer, sheet_name="Errores", index=False)
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
    os.replace(tmp, path)
    return df_estad, df_err


def main():
    parser = argparse.ArgumentParser(description="Segmentación robusta ")
    parser.add_argument("--casos", default=CASOS_XLSX, help="Excel del paso 1 v2")
    parser.add_argument("--salida", default=SALIDA_XLSX)
    parser.add_argument("--base-segmentacion", default=BASE_SEG_XLSX)
    parser.add_argument("--modo", choices=["incompletos", "sin_registro", "pacientes", "todos", "verificar"], default=None)
    parser.add_argument("--forzar", action="store_true", help="Reprocesar aunque ya exista el archivo de segmentación")
    parser.add_argument("--checkpoint", type=int, default=5, help="Guardar Excel cada N casos procesados (0 = solo al final)")
    parser.add_argument("--pacientes", nargs="*", default=[])
    args = parser.parse_args()

    if not os.path.exists(args.casos):
        raise FileNotFoundError(
            f"No existe {args.casos}. Ejecuta primero 1_filtrado_estadistica.py."
        )
    casos_df = pd.read_excel(args.casos, sheet_name="Casos Válidos")
    casos_df = casos_df.drop_duplicates(["Paciente", "Carpeta"], keep="last")

    ensure_output(args.base_segmentacion, args.salida)
    df_estad = cargar_hoja(args.salida, "Estadísticas", COLUMNAS)
    df_err = cargar_hoja(args.salida, "Errores", COLUMNAS_ERRORES)
    df_estad.drop_duplicates(["Paciente", "Carpeta"], keep="last", inplace=True)
    df_err.drop_duplicates(["Paciente", "Carpeta"], keep="last", inplace=True)

    modo = args.modo
    if modo is None:
        print("\n=== SEGMENTACIÓN  ===")
        print("1. Procesar solo archivos de segmentación faltantes")
        print("2. Reparar válidos sin registro estadístico")
        print("3. Procesar pacientes concretos")
        print("4. Verificar archivos + registros")
        print("5. Procesar todos")
        print("6. Salir")
        op = input("Opción: ").strip()
        modo = {"1": "incompletos", "2": "sin_registro", "3": "pacientes", "4": "verificar", "5": "todos"}.get(op)
        if op == "6":
            return
        if modo is None:
            raise SystemExit("Opción inválida")

    valid_pairs = {clave_par(p, c) for p, c in zip(casos_df["Paciente"], casos_df["Carpeta"])}

    # Limpia del Excel de segmentación pares que ya no pertenecen a Casos Válidos.
    before_est, before_err = len(df_estad), len(df_err)
    if not df_estad.empty:
        df_estad = df_estad.loc[[clave_par(p, c) in valid_pairs for p, c in zip(df_estad["Paciente"], df_estad["Carpeta"])]].copy()
    if not df_err.empty:
        df_err = df_err.loc[[clave_par(p, c) in valid_pairs for p, c in zip(df_err["Paciente"], df_err["Carpeta"])]].copy()
    limpiados_est = before_est - len(df_estad)
    limpiados_err = before_err - len(df_err)

    missing_file_mask = ~casos_df.apply(lambda r: os.path.exists(seg_path_for(r)), axis=1)
    incompletos = casos_df.loc[missing_file_mask].copy()

    registrados = pares_df(df_estad)
    missing_reg_mask = [clave_par(p, c) not in registrados for p, c in zip(casos_df["Paciente"], casos_df["Carpeta"])]
    sin_registro = casos_df.loc[missing_reg_mask].copy()

    if modo == "verificar":
        print(f"Casos válidos sin archivo de segmentación: {len(incompletos)}")
        print(f"Casos válidos sin registro estadístico: {len(sin_registro)}")
        print(f"Registros estadísticos retirados por ya no ser válidos: {limpiados_est}")
        print(f"Errores retirados por ya no ser válidos: {limpiados_err}")
        segmentables = []
        volver_paso1 = []
        rutas_no_resueltas = []
        for _, r in incompletos.iterrows():
            p, c = _norm(r["Paciente"]), _norm(r["Carpeta"])
            ruta = resolver_ruta_serie(r)
            if ruta is None:
                rutas_no_resueltas.append((p, c, "RUTA_NO_RESUELTA"))
                continue
            try:
                aud = auditar_cabeceras_rapido(str(ruta), max_muestras=7)
                estado = aud.get("Estado lectura", "")
                if estado in {"CT_APARENTA_SEGMENTABLE", "CT_APARENTA_SEGMENTABLE_MIXTA"}:
                    segmentables.append((p, c))
                elif estado in {"CT_NO_HU_MUESTRA", "MODALIDAD_NO_CT_MUESTRA"}:
                    volver_paso1.append((p, c, estado))
                else:
                    rutas_no_resueltas.append((p, c, estado or "REVISAR"))
            except Exception as e:
                rutas_no_resueltas.append((p, c, f"ERROR_AUDITORIA: {e}"))

        print(f"  Segmentables confirmadas: {len(segmentables)}")
        print(f"  Deben volver al paso 1: {len(volver_paso1)}")
        print(f"  Ruta no resuelta: {len(rutas_no_resueltas)}")
        if volver_paso1:
            print("\nSeries que NO deben segmentarse hasta reclasificarlas en paso 1:")
            for p, c, code in volver_paso1:
                print(f" - {p} / {c} -> {code}")
        if rutas_no_resueltas:
            print("\nRutas no resueltas (no se cambia su clasificación):")
            for p, c, code in rutas_no_resueltas:
                print(f" - {p} / {c} -> {code}")
        if len(sin_registro):
            print("\nHay segmentaciones cuyo archivo existe pero falta su fila en Estadísticas.")
            print("Ejecuta: python 2_segmentacion.py --modo sin_registro")
        return
    elif modo == "incompletos":
        work = incompletos
    elif modo == "sin_registro":
        work = sin_registro
    elif modo == "pacientes":
        pats = {p.strip() for p in args.pacientes if p.strip()}
        if not pats:
            raw = input("Pacientes separados por coma: ").strip()
            pats = {p.strip() for p in raw.split(",") if p.strip()}
        work = casos_df.loc[casos_df["Paciente"].map(_norm).isin(pats)].copy()
    else:
        work = casos_df.copy()

    print(f"Casos a procesar: {len(work)}")
    if limpiados_est or limpiados_err:
        print(f"Limpieza de registros ya no válidos: Estadísticas={limpiados_est}, Errores={limpiados_err}")
    ok = err = 0
    for n, (_, row) in enumerate(work.iterrows(), 1):
        p = estandarizar_clave(row["Paciente"])
        c = estandarizar_clave(row["Carpeta"])
        print(f"[{n}/{len(work)}] {p}/{c}")

        # incompletos/pacientes omiten archivos existentes. sin_registro los REPROCESA
        # expresamente para reconstruir la fila estadística perdida tras una interrupción.
        if not args.forzar and modo not in {"todos", "sin_registro"} and os.path.exists(seg_path_for(row)):
            print("  🟢 Ya existe segmentación; se omite.")
            continue

        status, result, audit = segmentar_caso(row)
        df_estad = quitar_par(df_estad, p, c)
        df_err = quitar_par(df_err, p, c)

        if status == "ok":
            df_estad = pd.concat([df_estad, pd.DataFrame([result])], ignore_index=True, sort=False)
            ok += 1
            print(f"  ✅ Segmentado: {result['Slices']} cortes | {result.get('Estado lectura','')}")
            log_evento(f"OK {p}/{c}: {result.get('Estado lectura','')}")
        else:
            df_err = pd.concat([df_err, pd.DataFrame([result])], ignore_index=True, sort=False)
            err += 1
            print(f"  ❌ {result['Codigo']}: {result['Motivo']}")
            log_evento(f"ERROR {p}/{c}: {result['Codigo']} - {result['Motivo']}")

        if args.checkpoint and n % args.checkpoint == 0:
            df_estad, df_err = guardar_estado_excel(args.salida, df_estad, df_err, ok, err)
            print(f"  💾 Checkpoint guardado ({n}/{len(work)})")

    df_estad.drop_duplicates(["Paciente", "Carpeta"], keep="last", inplace=True)
    df_err.drop_duplicates(["Paciente", "Carpeta"], keep="last", inplace=True)

    df_estad, df_err = guardar_estado_excel(args.salida, df_estad, df_err, ok, err)

    print("\n===== FINAL =====")
    print(f"OK esta corrida: {ok}")
    print(f"Errores esta corrida: {err}")
    print(f"Archivo: {args.salida}")


if __name__ == "__main__":
    main()
