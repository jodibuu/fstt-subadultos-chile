# Mapa entre la memoria y el repositorio

**Memoria de título (Diéguez-Burgos, 2026)**: *Medición del espesor tisular facial en subadultos chilenos mediante tomografía computarizada: Un enfoque semiautomatizado para la aproximación facial forense*.

| Capítulo/etapa de la memoria | Anexo | Archivo |
|---|---|---|
| §6.2.1 Ética y autorizaciones | A | Solo en la memoria, no en el repositorio |
| §6.3.2 Codificación y directorios | B | `codigo/config_pipeline.py` |
| §6.3.3 Validación de series DICOM | C | `codigo/1_filtrado_estadistica.py`, `codigo/dicom_reader_robusto.py` |
| §6.3.4–6.3.8 Segmentación | D | `codigo/2_segmentacion.py` |
| §6.4 Control de calidad visual | E | `codigo/3_renders.py` |
| §6.5.1–6.5.2 Extracción y suavizado | F | `codigo/4_mallasply.py` |
| §6.5.3 Convergencia | G | `codigo/analisisconvergencia.py` |
| §6.5.4 Orientación | H | `codigo/preprocesar_mallas.py` |
| §6.6 Landmarks, FSTT y CCU | I | `blender/tesis_landmarks.py` |
| §6.8 Análisis estadístico y figuras | J | `codigo/estadisticas_tesis_HC3.py`, `codigo/boxplots_fstt_por_grupo.py` |

**Advertencia:** la enumeración de anexos del paquete ZIP preliminar estaba desplazada, ya que no incluía el Anexo A de ética. Esta versión corrige la correspondencia con la memoria sin cambiar los nombres de los scripts.
