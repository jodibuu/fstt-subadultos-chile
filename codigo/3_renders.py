"""Generación de vistas estandarizadas para control de calidad.

Produce seis vistas de cada una de las dos máscaras segmentadas y registra
las rutas de salida en una planilla de control.
"""
import os
import argparse
import shutil
from datetime import datetime
import time
import numpy as np
import pandas as pd
import SimpleITK as sitk
from vedo import Volume, Plotter
from PIL import Image

from config_pipeline import RESULTADOS_FOLDER, SEGMENT_FOLDER
DEFAULT_EXCEL_BASE = os.path.join(str(RESULTADOS_FOLDER), "estadisticas_segmentacion_v2.xlsx")
DEFAULT_EXCEL_CON = DEFAULT_EXCEL_BASE.replace(".xlsx", "_con_renders.xlsx")
OUTPUT_FOLDER = os.path.join(str(RESULTADOS_FOLDER), "renders_segmentacion")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

VISTAS = {
    "frontal":         {"offset": (0, -3.5, 0),     "viewup": (0, 0, 1),  "rotar": 90},
    "superior":        {"offset": (3.5, -1.5, 0),   "viewup": (0, -1, 0), "rotar": -180},
    "inferior":        {"offset": (-3.5, -1.5, 0),  "viewup": (0, 1, 0),  "rotar": 180},
    "posterior":       {"offset": (0, 3.5, 0),      "viewup": (0, 0, 1),  "rotar": 270},
    "perfil_derecho":  {"offset": (0, 0, 3.5),      "viewup": (0, 1, 0),  "rotar": 90},
    "perfil_izq":      {"offset": (0, 0, -3.5),     "viewup": (0, 1, 0),  "rotar": 270}
}
RENDER_COLS = [f"PNG_{t}_{v}" for t in ("Blando", "Hueso") for v in VISTAS.keys()]


def choose_excel(path_arg: str | None) -> str:
    if path_arg:
        return path_arg
    if not os.path.exists(DEFAULT_EXCEL_CON):
        if not os.path.exists(DEFAULT_EXCEL_BASE):
            raise FileNotFoundError(f"No existe la planilla de segmentación: {DEFAULT_EXCEL_BASE}")
        shutil.copy2(DEFAULT_EXCEL_BASE, DEFAULT_EXCEL_CON)
    return DEFAULT_EXCEL_CON


def ensure_columns(df: pd.DataFrame):
    for c in RENDER_COLS:
        if c not in df.columns:
            df[c] = ""
    if "Spacing" not in df.columns:
        df["Spacing"] = ""


def png_exists(val) -> bool:
    if pd.isna(val):
        return False
    s = str(val).strip()
    return bool(s) and os.path.exists(s)


def get_expected_png(paciente_id: str, carpeta_id: str, tipo: str, vista_name: str) -> str:
    return os.path.join(OUTPUT_FOLDER, f"{paciente_id}_{carpeta_id}_{tipo}_{vista_name}.png")


def backup_excel(excel_path: str) -> str:
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    stem, ext = os.path.splitext(excel_path)
    out = f"{stem}_BACKUP_RENDERS_{ts}{ext}"
    shutil.copy2(excel_path, out)
    return out


def get_pos(center, offset, scale):
    return tuple(center[i] + offset[i] * scale for i in range(3))


def render_vistas_tipo(vol, paciente_id, carpeta_id, tipo, idx_df, df, vistas_a_hacer):
    if not vistas_a_hacer:
        return 0
    p = Plotter(offscreen=True, size=(900, 900), bg="white")  # type: ignore[arg-type]
    p.show(vol, axes=0, resetcam=True)
    cam = p.camera
    assert cam is not None
    center = vol.center_of_mass()
    dist = max(vol.bounds()[1::2]) * 1.5
    count = 0
    for vista_name in vistas_a_hacer:
        info = VISTAS[vista_name]
        cam.SetPosition(get_pos(center, info["offset"], dist))
        cam.SetFocalPoint(center)
        cam.SetViewUp(info["viewup"])
        cam.Zoom(1.3)
        filename = get_expected_png(paciente_id, carpeta_id, tipo, vista_name)
        p.screenshot(filename)
        rotar = info.get("rotar", 0)
        if rotar:
            img = Image.open(filename)
            img = img.rotate(rotar, expand=True)
            img.save(filename)
        df.at[idx_df, f"PNG_{tipo}_{vista_name}"] = filename
        count += 1
    p.close()
    return count


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default=None, help="Excel de estadísticas/renders")
    ap.add_argument("--pacientes", nargs="*", default=None, help="Lista de pacientes a procesar")
    ap.add_argument("--forzar", action="store_true", help="Rehacer vistas aunque el PNG exista")
    ap.add_argument("--verbose", action="store_true")
    ap.add_argument("--log-every", type=int, default=25)
    args = ap.parse_args()

    excel = choose_excel(args.excel)
    df = pd.read_excel(excel, sheet_name="Estadísticas")
    ensure_columns(df)

    if args.pacientes:
        wanted = {p.strip() for p in args.pacientes if str(p).strip()}
        s_p = df["Paciente"].astype(str).str.strip() if "Paciente" in df.columns else pd.Series([], dtype=str)
        mask = s_p.isin(wanted)
        df_idx = list(df.index[mask])
    else:
        df_idx = list(df.index)

    print(f"\n===== RENDERS DE CONTROL =====")
    print(f"Archivo en uso: {excel}")
    if args.pacientes:
        print("Pacientes filtrados: " + ", ".join(args.pacientes))
    print(f"Filas candidatas: {len(df_idx)}")

    errores, completados = [], []
    done_cases = skip_cases = fail_cases = 0
    total_png = 0
    start = time.time()

    for n, idx in enumerate(df_idx, start=1):
        row = df.loc[idx]
        paciente_id = str(row.get("Paciente", "")).strip()
        carpeta_id  = str(row.get("Carpeta", "")).strip()
        grupo = str(row.get("Grupo", "")).strip()
        sexo  = str(row.get("Sexo", "")).strip()
        nombre_base = f"{paciente_id}_{carpeta_id}"
        out_dir = os.path.join(SEGMENT_FOLDER, grupo, sexo, paciente_id, nombre_base)
        ruta_label = os.path.join(out_dir, f"{nombre_base}_segmentacion.label.nii.gz")

        if not os.path.exists(ruta_label):
            errores.append({"Paciente": paciente_id, "Carpeta": carpeta_id, "Motivo": "No se encontró segmentación"})
            fail_cases += 1
            if args.verbose:
                print(f"⚠️ {paciente_id}/{carpeta_id}: sin segmentación")
            continue

        faltantes = []
        for t in ("Blando", "Hueso"):
            for v in VISTAS.keys():
                col = f"PNG_{t}_{v}"
                expected = str(row.get(col, "")).strip() if not pd.isna(row.get(col, "")) else ""
                if args.forzar:
                    faltantes.append((t, v))
                    continue
                if expected and os.path.exists(expected):
                    # ya existe donde dice el Excel
                    continue
                # si la celda está vacía o apunta a algo inexistente, intenta el nombre esperado estándar
                if os.path.exists(get_expected_png(paciente_id, carpeta_id, t, v)):
                    df.at[idx, col] = get_expected_png(paciente_id, carpeta_id, t, v)
                    continue
                faltantes.append((t, v))

        if not faltantes:
            completados.append({"Paciente": paciente_id, "Carpeta": carpeta_id, "Motivo": "PNG existentes"})
            skip_cases += 1
            if args.verbose:
                print(f"🟢 Skip {paciente_id}/{carpeta_id} (renders ya existen)")
            continue

        try:
            img = sitk.ReadImage(ruta_label)
            img = sitk.DICOMOrient(img, "LPS")
            seg = sitk.GetArrayFromImage(img)
            spacing = img.GetSpacing()  # x,y,z or preserved metadata
            scale = (spacing[2], spacing[1], spacing[0])
            df.at[idx, "Spacing"] = f"({spacing[2]:.4f}, {spacing[1]:.4f}, {spacing[0]:.4f})"
            estructuras = {
                "Blando": (seg == 1).astype(np.uint8),
                "Hueso":  (seg == 2).astype(np.uint8),
            }
            rendered = 0
            for tipo, array in estructuras.items():
                vistas_tipo = [v for (t, v) in faltantes if t == tipo]
                if not vistas_tipo:
                    continue
                vol = Volume(array, spacing=scale).alpha([0, 1]).cmap("gray" if tipo == "Blando" else "bone")
                rendered += render_vistas_tipo(vol, paciente_id, carpeta_id, tipo, idx, df, vistas_tipo)
            if rendered > 0:
                done_cases += 1
                total_png += rendered
                if args.verbose:
                    print(f"✅ {paciente_id}/{carpeta_id}: {rendered} PNG")
            else:
                skip_cases += 1
        except Exception as e:
            errores.append({"Paciente": paciente_id, "Carpeta": carpeta_id, "Motivo": str(e)})
            fail_cases += 1
            print(f"❌ Error {paciente_id}/{carpeta_id}: {e}")

        if (n % args.log_every == 0) and not args.verbose:
            print(f"… progreso: {n}/{len(df_idx)} | casos con nuevos renders {done_cases}, omitidos {skip_cases}, errores {fail_cases}")

    bak = backup_excel(excel)
    with pd.ExcelWriter(excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="Estadísticas", index=False)
        pd.DataFrame(errores, columns=["Paciente", "Carpeta", "Motivo"]).to_excel(
            writer, sheet_name="Errores_Renders", index=False)
        pd.DataFrame(completados, columns=["Paciente", "Carpeta", "Motivo"]).to_excel(
            writer, sheet_name="Completados_Renders", index=False)

    elapsed = time.time() - start
    print("\n===== FINAL =====")
    print(f"Casos considerados: {len(df_idx)}")
    print(f"Casos con nuevos renders: {done_cases}")
    print(f"Casos omitidos (ya completos): {skip_cases}")
    print(f"Errores: {fail_cases}")
    print(f"PNG generados: {total_png}")
    print(f"Backup Excel: {bak}")
    print(f"Excel actualizado: {excel}")
    print(f"Tiempo: {elapsed/60:.1f} min")

if __name__ == "__main__":
    main()
