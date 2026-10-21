"""
analisisconvergencia.py — Análisis de convergencia del suavizado laplaciano.

Para dos mallas de referencia (tejido óseo y tejido blando),
aplica suavizado laplaciano incremental de 0 a 50 iteraciones y calcula cinco
métricas de convergencia por iteración:
    1. Distancia media entre vértices respecto al original (mm).
    2. Distancia máxima entre vértices respecto al original (mm).
    3. Ángulo medio entre normales de vértice respecto al original (°).
    4. Distancia media entre centroides de caras (mm).
    5. Ángulo medio entre normales de cara (°).

Cada malla intermedia se exporta en PLY con colores viridis proporcionales a
la distancia al original. Los resultados numéricos se guardan en Excel.

Uso:
    python analisisconvergencia.py

Dependencias:
    open3d, numpy, pandas, matplotlib, xlsxwriter
"""

import open3d as o3d
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
import os

from config_pipeline import MALLAS_FOLDER, RESULTADOS_FOLDER

# === CONFIGURACIÓN ===
MODELOS = {
    "duro": os.environ.get("FSTT_MALLA_REFERENCIA_HUESO", str(MALLAS_FOLDER / "caso_referencia_hueso.ply")),
    "blando": os.environ.get("FSTT_MALLA_REFERENCIA_BLANDO", str(MALLAS_FOLDER / "caso_referencia_blando.ply")),
}
N_ITER = 50
OUTPUT_DIR = Path(os.environ.get("FSTT_CONVERGENCIA_SALIDA", str(RESULTADOS_FOLDER / "convergencia")))
(OUTPUT_DIR / "ply").mkdir(parents=True, exist_ok=True)
(OUTPUT_DIR / "xlsx").mkdir(exist_ok=True)

# === FUNCIONES ===

def calcular_normales(malla):
    """Recalcula normales de vértices y caras de una malla Open3D."""
    malla.compute_vertex_normals()
    malla.compute_triangle_normals()

def suavizar_malla(malla, n_iter):
    """Aplica n_iter iteraciones de suavizado laplaciano y retorna la malla suavizada."""
    return malla.filter_smooth_laplacian(number_of_iterations=n_iter)

def distancia_por_vertice(v1, v2):
    """Distancia euclidiana por vértice entre dos arrays (N, 3)."""
    return np.linalg.norm(v1 - v2, axis=1)

def angulo_entre_normales(n1, n2):
    """Ángulo en grados entre pares de vectores normales."""
    dot = np.einsum('ij,ij->i', n1, n2)
    dot = np.clip(dot, -1.0, 1.0)
    ang = np.arccos(dot) * 180 / np.pi
    return ang

def centroides_caras(mesh):
    """Calcula los centroides de todas las caras triangulares de la malla."""
    verts = np.asarray(mesh.vertices)
    tris = np.asarray(mesh.triangles)
    return np.mean(verts[tris], axis=1)

def aplicar_color_por_distancia(mesh, distancias):
    """Asigna colores viridis a los vértices según distancia normalizada."""
    d_norm = (distancias - distancias.min()) / (distancias.max() - distancias.min() + 1e-8)
    cmap = plt.get_cmap("viridis")
    colores = cmap(d_norm)[:, :3]  # RGB sin alfa
    mesh.vertex_colors = o3d.utility.Vector3dVector(colores)

# === PROCESO PRINCIPAL ===

def procesar_modelo(nombre, ruta):
    """Aplica suavizado incremental (0-50 iter) y calcula métricas de convergencia."""
    print(f" Procesando: {nombre}")
    malla_original = o3d.io.read_triangle_mesh(ruta)
    calcular_normales(malla_original)

    v0 = np.asarray(malla_original.vertices)
    n0 = np.asarray(malla_original.vertex_normals)
    f0 = np.asarray(malla_original.triangles)
    nt0 = np.asarray(malla_original.triangle_normals)
    c0 = centroides_caras(malla_original)

    resultados = []

    for i in range(N_ITER + 1):
        if i == 0:
            malla = malla_original
        else:
            malla = suavizar_malla(malla_original, i)
            calcular_normales(malla)

        v_i = np.asarray(malla.vertices)
        n_i = np.asarray(malla.vertex_normals)
        c_i = centroides_caras(malla)
        nt_i = np.asarray(malla.triangle_normals)

        dist_vert = distancia_por_vertice(v0, v_i)
        ang_vert = angulo_entre_normales(n0, n_i)

        dist_cara = distancia_por_vertice(c0, c_i)
        ang_cara = angulo_entre_normales(nt0, nt_i)

        # Guardar métricas
        resultados.append({
            "iter": i,
            "mean_vertex_dist": dist_vert.mean(),
            "max_vertex_dist": dist_vert.max(),
            "mean_vertex_angle": ang_vert.mean(),
            "mean_face_dist": dist_cara.mean(),
            "mean_face_angle": ang_cara.mean()
        })

        # === Guardar modelo .ply con colores ===
        aplicar_color_por_distancia(malla, dist_vert)
        ply_path = OUTPUT_DIR / "ply" / f"{nombre}_iter_{i:02}.ply"
        o3d.io.write_triangle_mesh(str(ply_path), malla)

    # === Guardar resultados .xlsx ===
    df = pd.DataFrame(resultados)
    xlsx_path = OUTPUT_DIR / "xlsx" / f"{nombre}_resultados.xlsx"
    with pd.ExcelWriter(xlsx_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=nombre)
        worksheet = writer.sheets[nombre]
        for idx, col in enumerate(df.columns):
            worksheet.set_column(idx, idx, 20)

    print(f" {nombre} procesado correctamente.\n")

# === CORRER ===

if __name__ == "__main__":
    for nombre, ruta in MODELOS.items():
        procesar_modelo(nombre, ruta)

