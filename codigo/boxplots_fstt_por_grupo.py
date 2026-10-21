from pathlib import Path
import os
from config_pipeline import RESULTADOS_FOLDER
import math
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

ARCHIVO_EXCEL = Path(os.environ.get("FSTT_BOXPLOTS_EXCEL", str(RESULTADOS_FOLDER / "RESULTADOS_HC3.xlsx"))).expanduser()
HOJA = "Base_principal"

CARPETA_SALIDA = Path(os.environ.get("FSTT_BOXPLOTS_DIR", str(RESULTADOS_FOLDER / "figuras_boxplots_fstt_final"))).expanduser()
CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

plt.rcParams.update({
    "font.family": "serif",
    "font.serif": ["Times New Roman", "Times", "DejaVu Serif"],
    "font.size": 10,
    "axes.titlesize": 12.5,
    "axes.labelsize": 11,
    "xtick.labelsize": 8.5,
    "ytick.labelsize": 9.5,
    "legend.fontsize": 9,
    "figure.dpi": 300,
    "savefig.dpi": 300,
    "pdf.fonttype": 42,
    "ps.fonttype": 42
})

# =========================================================
# GRUPOS Y COLORES
# =========================================================

GRUPOS = ["F_6-11", "F_12-18", "M_6-11", "M_12-18"]

ETIQUETAS_GRUPOS = {
    "F_6-11": "Femenino 6–11",
    "F_12-18": "Femenino 12–18",
    "M_6-11": "Masculino 6–11",
    "M_12-18": "Masculino 12–18",
}

# Paleta coherente con tus otras figuras
COLORES = {
    "F_6-11": "#f4a9c8",   # rosa claro
    "F_12-18": "#c91872",  # fucsia
    "M_6-11": "#8bcdeb",   # celeste
    "M_12-18": "#2457a6",  # azul
}

# =========================================================
# LANDMARKS Y ORDEN
# =========================================================

PANEL_A = ["g", "n", "mn", "rhi", "sn", "mp", "ls", "li", "mls", "pg", "gn", "m"]

PANEL_B = ["mso-izq", "mso-der", "mio-izq", "mio-der", "acp-izq", "acp-der"]

PANEL_C = ["zy-izq", "zy-der", "mr-izq", "mr-der", "go-izq", "go-der", "mmb-izq", "mmb-der"]

TODOS = PANEL_A + PANEL_B + PANEL_C

ETIQUETAS_LM = {
    "g": "G",
    "n": "N",
    "mn": "MN",
    "rhi": "RHI",
    "sn": "SN",
    "mp": "MP",
    "ls": "LS",
    "li": "LI",
    "mls": "MLS",
    "pg": "PG",
    "gn": "GN",
    "m": "M",
    "mso-izq": "MSO-I",
    "mso-der": "MSO-D",
    "mio-izq": "MIO-I",
    "mio-der": "MIO-D",
    "acp-izq": "ACP-I",
    "acp-der": "ACP-D",
    "zy-izq": "ZY-I",
    "zy-der": "ZY-D",
    "mr-izq": "MR-I",
    "mr-der": "MR-D",
    "go-izq": "GO-I",
    "go-der": "GO-D",
    "mmb-izq": "MMB-I",
    "mmb-der": "MMB-D",
}

# =========================================================
# CARGA Y PREPARACIÓN DE DATOS
# =========================================================

df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA)

df["Status"] = df["Status"].astype(str).str.upper().str.strip()
df["Landmark"] = df["Landmark"].astype(str).str.lower().str.strip()
df["Grupo"] = df["Grupo"].astype(str).str.strip()

df = df[
    (df["Valido"] == True) &
    (df["Status"] == "SET") &
    (df["Grupo"].isin(GRUPOS)) &
    (df["Landmark"].isin(TODOS))
].copy()

# =========================================================
# FUNCIONES AUXILIARES
# =========================================================

def limite_superior_bonito(valores):
    """
    Calcula un límite superior de eje Y 'bonito'
    para que no queden cortadas las cajas.
    """
    if len(valores) == 0:
        return 10
    maximo = np.nanmax(valores)
    if maximo <= 10:
        paso = 1
    elif maximo <= 20:
        paso = 2
    else:
        paso = 5
    return math.ceil((maximo + 0.8) / paso) * paso


def etiqueta_con_n(df_sub, lm):
    """
    Genera la etiqueta del landmark con los n debajo.
    Formato:
    LM
    12|15|14|16
    """
    ns = []
    for grupo in GRUPOS:
        n = df_sub.loc[
            (df_sub["Landmark"] == lm) & (df_sub["Grupo"] == grupo),
            "Thickness_mm"
        ].notna().sum()
        ns.append(str(n))
    return f"{ETIQUETAS_LM[lm]}\n" + "|".join(ns)


def preparar_datos_panel(df_sub, landmarks):
    """
    Devuelve:
    - datos por grupo
    - n por grupo y landmark implícito en las etiquetas
    - valores completos para calcular el eje Y
    """
    datos_grupo = {g: [] for g in GRUPOS}
    todos_valores = []

    for lm in landmarks:
        for g in GRUPOS:
            vals = df_sub.loc[
                (df_sub["Landmark"] == lm) & (df_sub["Grupo"] == g),
                "Thickness_mm"
            ].dropna().values
            datos_grupo[g].append(vals)
            todos_valores.extend(vals.tolist())

    return datos_grupo, np.array(todos_valores)


def graficar_panel(ax, df_sub, landmarks, titulo_panel, y_max=None):
    """
    Dibuja un panel de boxplots verticales.
    """
    datos_grupo, todos_valores = preparar_datos_panel(df_sub, landmarks)

    if y_max is None:
        y_max = limite_superior_bonito(todos_valores)

    centros = np.arange(1, len(landmarks) + 1)
    offsets = [-0.27, -0.09, 0.09, 0.27]
    ancho = 0.16

    for i, grupo in enumerate(GRUPOS):
        posiciones = centros + offsets[i]
        bp = ax.boxplot(
            datos_grupo[grupo],
            positions=posiciones,
            widths=ancho,
            patch_artist=True,
            manage_ticks=False,
            showfliers=True,
            whis=1.5,
            medianprops=dict(color="#1f1f1f", linewidth=1.2),
            whiskerprops=dict(color="#555555", linewidth=0.8),
            capprops=dict(color="#555555", linewidth=0.8),
            boxprops=dict(edgecolor="#4a4a4a", linewidth=0.8),
            flierprops=dict(
                marker="o",
                markersize=2.3,
                markerfacecolor="#666666",
                markeredgecolor="#666666",
                alpha=0.30
            )
        )

        for box in bp["boxes"]:
            box.set_facecolor(COLORES[grupo])
            box.set_alpha(0.90)

    etiquetas_x = [etiqueta_con_n(df_sub, lm) for lm in landmarks]

    ax.set_xticks(centros)
    ax.set_xticklabels(etiquetas_x, rotation=0, ha="center")
    ax.set_ylim(0, y_max)
    ax.set_ylabel("Espesor tisular facial (mm)")
    ax.set_title(titulo_panel, loc="left", pad=10)

    ax.grid(axis="y", linewidth=0.5, alpha=0.22)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def guardar_figura(fig, nombre_base):
    """
    Guarda PNG, PDF y SVG.
    """
    for ext in ["png", "pdf", "svg"]:
        fig.savefig(
            CARPETA_SALIDA / f"{nombre_base}.{ext}",
            bbox_inches="tight",
            facecolor="white"
        )

# =========================================================
# LÍMITE GLOBAL PARA COMPARABILIDAD
# =========================================================

y_max_global = limite_superior_bonito(df["Thickness_mm"].dropna().values)

# =========================================================
# FIGURA PRINCIPAL EN 3 PANELES
# =========================================================

fig, axes = plt.subplots(
    nrows=3,
    ncols=1,
    figsize=(13, 18),
    sharey=True
)

graficar_panel(axes[0], df, PANEL_A, "A. Landmarks de línea media", y_max=y_max_global)
graficar_panel(axes[1], df, PANEL_B, "B. Landmarks orbitarios y nasales bilaterales", y_max=y_max_global)
graficar_panel(axes[2], df, PANEL_C, "C. Landmarks laterales y mandibulares bilaterales", y_max=y_max_global)

# Leyenda general
handles = [
    Patch(facecolor=COLORES[g], edgecolor="#4a4a4a", label=ETIQUETAS_GRUPOS[g])
    for g in GRUPOS
]

fig.legend(
    handles=handles,
    loc="upper center",
    bbox_to_anchor=(0.5, 0.995),
    ncol=4,
    frameon=False
)

fig.suptitle(
    "Distribución del espesor tisular facial por landmark según grupo sexo-etario",
    y=0.999,
    fontsize=14
)

fig.text(
    0.5,
    0.01,
    "Bajo cada landmark se indica el tamaño muestral disponible en el orden "
    "F 6–11 | F 12–18 | M 6–11 | M 12–18. "
    "Las cajas representan Q1–Q3; la línea central corresponde a la mediana; "
    "los bigotes se extienden hasta 1,5 veces el rango intercuartílico.",
    ha="center",
    va="bottom",
    fontsize=8.7
)

plt.tight_layout(rect=[0.03, 0.04, 0.97, 0.97])
guardar_figura(fig, "Figura_principal_boxplots_3_paneles")
plt.close(fig)

# =========================================================
# FIGURAS INDIVIDUALES POR PANEL
# =========================================================

def figura_individual(landmarks, titulo, nombre_base, alto=6.5, ancho=12):
    fig, ax = plt.subplots(figsize=(ancho, alto))
    graficar_panel(ax, df, landmarks, titulo, y_max=y_max_global)

    handles = [
        Patch(facecolor=COLORES[g], edgecolor="#4a4a4a", label=ETIQUETAS_GRUPOS[g])
        for g in GRUPOS
    ]
    fig.legend(
        handles=handles,
        loc="upper center",
        bbox_to_anchor=(0.5, 0.98),
        ncol=4,
        frameon=False
    )

    fig.text(
        0.5,
        0.01,
        "n bajo cada landmark en el orden F 6–11 | F 12–18 | M 6–11 | M 12–18.",
        ha="center",
        va="bottom",
        fontsize=8.5
    )

    plt.tight_layout(rect=[0.03, 0.04, 0.97, 0.93])
    guardar_figura(fig, nombre_base)
    plt.close(fig)

figura_individual(
    PANEL_A,
    "A. Landmarks de línea media",
    "Figura_A_linea_media",
    alto=6.6,
    ancho=13
)

figura_individual(
    PANEL_B,
    "B. Landmarks orbitarios y nasales bilaterales",
    "Figura_B_orbitarios_nasales",
    alto=6.6,
    ancho=10
)

figura_individual(
    PANEL_C,
    "C. Landmarks laterales y mandibulares bilaterales",
    "Figura_C_laterales_mandibulares",
    alto=6.6,
    ancho=11
)

# =========================================================
# FIGURA COMPLETA PARA ANEXO
# =========================================================

fig, ax = plt.subplots(figsize=(20, 8.5))
graficar_panel(
    ax,
    df,
    TODOS,
    "Distribución completa del espesor tisular facial por landmark y grupo sexo-etario",
    y_max=y_max_global
)

handles = [
    Patch(facecolor=COLORES[g], edgecolor="#4a4a4a", label=ETIQUETAS_GRUPOS[g])
    for g in GRUPOS
]

fig.legend(
    handles=handles,
    loc="upper center",
    bbox_to_anchor=(0.5, 0.98),
    ncol=4,
    frameon=False
)

fig.text(
    0.5,
    0.01,
    "Bajo cada landmark se indica el tamaño muestral disponible en el orden "
    "F 6–11 | F 12–18 | M 6–11 | M 12–18.",
    ha="center",
    va="bottom",
    fontsize=8.7
)

plt.tight_layout(rect=[0.03, 0.04, 0.97, 0.93])
guardar_figura(fig, "Figura_anexo_boxplots_26_landmarks")
plt.close(fig)

print(f"Listo. Figuras guardadas en: {CARPETA_SALIDA.resolve()}")