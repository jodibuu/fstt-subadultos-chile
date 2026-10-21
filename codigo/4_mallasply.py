"""Reconstrucción de mallas óseas y de tejido blando.

Extrae superficies con Marching Cubes (nivel 0,5), aplica 20 iteraciones de
suavizado laplaciano y exporta las mallas en formato PLY.
"""
from __future__ import annotations

import argparse
import os
import shutil
import time
from datetime import datetime
from pathlib import Path

import numpy as np
import open3d as o3d
import pandas as pd
import SimpleITK as sitk
from skimage import measure

from config_pipeline import RESULTADOS_FOLDER, SEGMENT_FOLDER, MALLAS_FOLDER

DEFAULT_EXCEL = RESULTADOS_FOLDER / "estadisticas_segmentacion_v2_con_renders.xlsx"
DEFAULT_MALLAS = MALLAS_FOLDER

EXCEL_COL_RUTA = "Ruta salida segmentación"
COL_MEJOR = "¿Es la mejor carpeta del paciente? (1 = Sí)"
COL_DESCARTAR = "¿Descartar? (1 = Sí)"
N_ITER = 20


def norm(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def cargar_segmentacion(path: Path):
    img = sitk.ReadImage(str(path))
    spacing_xyz = img.GetSpacing()
    array = sitk.GetArrayFromImage(img)
    return array, spacing_xyz


def generar_malla(array_bin, spacing_zyx):
    verts, faces, _, _ = measure.marching_cubes(
        array_bin, level=0.5, spacing=spacing_zyx
    )
    mesh = o3d.geometry.TriangleMesh()
    mesh.vertices = o3d.utility.Vector3dVector(verts)
    mesh.triangles = o3d.utility.Vector3iVector(faces)
    return mesh


def resolver_ruta_segmentacion(row) -> Path | None:
    """Resuelve la segmentación sin convertir una ruta histórica en un error real."""
    raw = norm(row.get(EXCEL_COL_RUTA, ""))
    candidatos: list[Path] = []

    def add(p):
        if not p:
            return
        p = Path(str(p).strip().strip('"').replace("\\", "/"))
        if str(p).lower() not in {str(x).lower() for x in candidatos}:
            candidatos.append(p)

    add(raw)

    if raw:
        normal = raw.replace("\\", "/")
        marker = "/2_PROCESAMIENTO/SEGMENTACIONES/"
        low = normal.lower()
        pos = low.find(marker.lower())
        if pos >= 0:
            suffix = normal[pos + len(marker):]
            add(SEGMENT_FOLDER / Path(suffix))

    grupo = norm(row.get("Grupo", ""))
    sexo = norm(row.get("Sexo", ""))
    paciente = norm(row.get("Paciente", ""))
    carpeta = norm(row.get("Carpeta", ""))
    if grupo and sexo and paciente and carpeta:
        nombre_base = f"{paciente}_{carpeta}"
        add(
            SEGMENT_FOLDER
            / grupo
            / sexo
            / paciente
            / nombre_base
            / f"{nombre_base}_segmentacion.label.nii.gz"
        )

    for p in candidatos:
        if p.is_file():
            return p
    return None


def cargar_excel(path: Path):
    xls = pd.ExcelFile(path)
    sheet = "Estadísticas" if "Estadísticas" in xls.sheet_names else xls.sheet_names[0]
    return pd.read_excel(path, sheet_name=sheet), sheet


def preparar_seleccion(df, pacientes=None, carpetas=None):
    if "Paciente" not in df.columns or "Carpeta" not in df.columns:
        raise KeyError("El Excel debe contener las columnas 'Paciente' y 'Carpeta'.")
    if EXCEL_COL_RUTA not in df.columns:
        raise KeyError(f"El Excel debe contener '{EXCEL_COL_RUTA}'.")

    pacientes = [str(x).strip() for x in (pacientes or []) if str(x).strip()]
    carpetas = [str(x).strip() for x in (carpetas or []) if str(x).strip()]

    # --carpetas es una selección explícita y tiene prioridad sobre los flags de QC.
    if carpetas:
        sel = df[df["Carpeta"].astype(str).isin(carpetas)].copy()
        faltan = sorted(set(carpetas) - set(sel["Carpeta"].astype(str)))
        if faltan:
            print("ADVERTENCIA: carpetas solicitadas no encontradas:")
            for x in faltan:
                print(f"  - {x}")
        if pacientes:
            sel = sel[sel["Paciente"].astype(str).isin(pacientes)].copy()
        return sel, "selección explícita por --carpetas"

    if COL_MEJOR not in df.columns:
        raise KeyError(
            f"El Excel no contiene '{COL_MEJOR}'. Usa un Excel de QC con la serie final "
            "marcada, o pasa --carpetas con la UID exacta que quieres convertir en malla."
        )

    mask = df[COL_MEJOR].fillna(0).eq(1)
    if COL_DESCARTAR in df.columns:
        mask &= ~df[COL_DESCARTAR].fillna(0).eq(1)
    mask &= df[EXCEL_COL_RUTA].notna()
    sel = df[mask].copy()

    if pacientes:
        sel = sel[sel["Paciente"].astype(str).isin(pacientes)].copy()
        encontrados = set(sel["Paciente"].astype(str))
        faltan = [p for p in pacientes if p not in encontrados]
        if faltan:
            print("ADVERTENCIA: pacientes pedidos sin una serie marcada como mejor y utilizable:")
            for p in faltan:
                candidatas = df[df["Paciente"].astype(str).eq(p)]
                print(f"  - {p}: {len(candidatas)} serie(s) en el Excel")
                for _, r in candidatas.iterrows():
                    print(
                        f"      {norm(r.get('Carpeta'))} | mejor={norm(r.get(COL_MEJOR))} "
                        f"| descartar={norm(r.get(COL_DESCARTAR))}"
                    )

    # Nunca generar dos series del mismo paciente accidentalmente.
    dup = sel[sel.duplicated(subset=["Paciente"], keep=False)]
    if not dup.empty:
        detalle = dup[["Paciente", "Carpeta"]].to_string(index=False)
        raise ValueError(
            "Hay pacientes con más de una serie seleccionada. Corrige el QC o usa --carpetas "
            "para indicar la serie exacta:\n" + detalle
        )

    return sel, "series marcadas como mejor en QC"


def backup_excel(excel: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    destino = excel.with_name(f"{excel.stem}_BACKUP_MALLAS_{stamp}{excel.suffix}")
    shutil.copy2(excel, destino)
    return destino


def main():
    parser = argparse.ArgumentParser(description="Genera mallas PLY desde segmentaciones seleccionadas.")
    parser.add_argument("--excel", default=str(DEFAULT_EXCEL), help="Excel con QC/selección de series.")
    parser.add_argument("--salida", default=str(DEFAULT_MALLAS), help="Carpeta base donde guardar las PLY.")
    parser.add_argument("--pacientes", nargs="*", default=[], help="Procesar solo estos pacientes.")
    parser.add_argument("--carpetas", nargs="*", default=[], help="Procesar solo estas carpetas/UID exactas.")
    parser.add_argument(
        "--forzar",
        action="store_true",
        help="Regenera y reemplaza PLY existentes. Sin esta opción, las existentes se omiten.",
    )
    args = parser.parse_args()

    excel = Path(args.excel)
    carpeta_base = Path(args.salida)
    if not excel.is_file():
        raise FileNotFoundError(f"No existe el Excel: {excel}")
    carpeta_base.mkdir(parents=True, exist_ok=True)

    df, sheet = cargar_excel(excel)
    seleccion, origen_sel = preparar_seleccion(df, args.pacientes, args.carpetas)

    if "Ruta malla blando" not in df.columns:
        df["Ruta malla blando"] = ""
    if "Ruta malla hueso" not in df.columns:
        df["Ruta malla hueso"] = ""

    print("===== RECONSTRUCCIÓN DE MALLAS =====")
    print(f"Excel: {excel}")
    print(f"Hoja: {sheet}")
    print(f"Selección: {origen_sel}")
    print(f"Casos seleccionados: {len(seleccion)}")
    print(f"Política PLY existente: {'REEMPLAZAR (--forzar)' if args.forzar else 'OMITIR'}")

    if seleccion.empty:
        print("No hay casos para procesar. No se modificó el Excel.")
        return

    inicio = time.time()
    generadas = 0
    omitidas = 0
    errores = 0
    sin_ruta = 0

    for n, (idx, row) in enumerate(seleccion.iterrows(), start=1):
        paciente = norm(row.get("Paciente"))
        carpeta = norm(row.get("Carpeta"))
        ruta = resolver_ruta_segmentacion(row)

        print(f"\n[{n}/{len(seleccion)}] {paciente} / {carpeta}")
        if ruta is None:
            print("  ERROR: no se encontró el NIfTI de segmentación.")
            sin_ruta += 1
            continue

        print(f"  Segmentación: {ruta}")
        try:
            array, spacing_xyz = cargar_segmentacion(ruta)
        except Exception as e:
            print(f"  ERROR al leer segmentación: {e}")
            errores += 1
            continue

        if array.ndim != 3:
            print(f"  ERROR: segmentación no 3D (shape={array.shape}).")
            errores += 1
            continue

        # SITK spacing = (x,y,z), array = (z,y,x).
        spacing_zyx = (spacing_xyz[2], spacing_xyz[1], spacing_xyz[0])
        nombre = f"{paciente}_{Path(ruta).stem.replace('_segmentacion.label', '')}"
        carpeta_paciente = carpeta_base / paciente
        carpeta_paciente.mkdir(parents=True, exist_ok=True)

        for tipo, valor in {"blando": 1, "hueso": 2}.items():
            salida = carpeta_paciente / f"{nombre}_{tipo}.ply"
            col = f"Ruta malla {tipo}"

            if salida.exists() and not args.forzar:
                print(f"  {tipo}: ya existe -> se omite")
                omitidas += 1
                df.at[idx, col] = str(salida.resolve())
                continue

            mask = (array == valor).astype(np.uint8)
            if int(mask.sum()) == 0:
                print(f"  ERROR {tipo}: máscara vacía")
                errores += 1
                continue

            try:
                mesh = generar_malla(mask, spacing_zyx)
                mesh = mesh.filter_smooth_laplacian(number_of_iterations=N_ITER)
                mesh.compute_vertex_normals()
                ok = o3d.io.write_triangle_mesh(str(salida), mesh)
                if not ok:
                    raise RuntimeError("Open3D devolvió False al escribir la PLY")
                df.at[idx, col] = str(salida.resolve())
                print(f"  {tipo}: guardada -> {salida}")
                generadas += 1
            except Exception as e:
                print(f"  ERROR {tipo}: {e}")
                errores += 1

    # Solo escribimos el Excel después de terminar la corrida.
    copia = backup_excel(excel)
    with pd.ExcelWriter(excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=sheet, index=False)

    dur = time.time() - inicio
    print("\n===== FINAL =====")
    print(f"Casos seleccionados: {len(seleccion)}")
    print(f"PLY generadas/reemplazadas: {generadas}")
    print(f"PLY existentes omitidas: {omitidas}")
    print(f"Errores de generación/lectura: {errores}")
    print(f"Segmentaciones no encontradas: {sin_ruta}")
    print(f"Backup Excel: {copia}")
    print(f"Excel actualizado: {excel}")
    print(f"Tiempo: {dur/60:.1f} min")


if __name__ == "__main__":
    main()
