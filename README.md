# Espesor tisular facial en subadultos: código de la memoria de título

**Repositorio de código y documentación complementaria de la memoria:**

> **Medición del espesor tisular facial en subadultos chilenos mediante tomografía computarizada: Un enfoque semiautomatizado para la aproximación facial forense**
>
> Josefa Diéguez-Burgos · Antropología Física · Universidad de Chile · 2026

Este repositorio reúne los programas desarrollados para organizar y validar series de tomografía computarizada (TC), segmentar superficies, preparar mallas tridimensionales, asistir la medición de **espesor tisular facial (FSTT)** en Blender y realizar los análisis estadísticos descritos en la memoria.

**Qué contiene:** código fuente, documentación metodológica y pruebas limitadas con archivos sintéticos. **Qué no contiene:** tomografías, mallas anatómicas individuales, planillas clínicas, mediciones de participantes, identificadores personales ni datos hospitalarios.

> **Alcance:** el flujo es **semiautomatizado**. La selección de la reconstrucción representativa, la revisión de segmentaciones y el posicionamiento/validación anatómica de los *landmarks* requieren intervención humana. Publicar el código no equivale a validar su exactitud en otros conjuntos de datos.

## Contenido y relación con los anexos

| Anexo de la memoria | Función | Archivo del repositorio |
|---|---|---|
| A | Aprobaciones éticas y administrativas | **Solo en la memoria**; no se publican los documentos clínicos/administrativos en este repositorio |
| B | Configuración de rutas | [`codigo/config_pipeline.py`](codigo/config_pipeline.py) |
| C | Validación de series DICOM | [`codigo/1_filtrado_estadistica.py`](codigo/1_filtrado_estadistica.py) y [`codigo/dicom_reader_robusto.py`](codigo/dicom_reader_robusto.py) |
| D | Segmentación de tejidos | [`codigo/2_segmentacion.py`](codigo/2_segmentacion.py) |
| E | Vistas de control visual | [`codigo/3_renders.py`](codigo/3_renders.py) |
| F | Extracción y suavizado de mallas | [`codigo/4_mallasply.py`](codigo/4_mallasply.py) |
| G | Análisis de convergencia del suavizado | [`codigo/analisisconvergencia.py`](codigo/analisisconvergencia.py) |
| H | Orientación geométrica | [`codigo/preprocesar_mallas.py`](codigo/preprocesar_mallas.py) |
| I | Complemento de medición en Blender | [`blender/tesis_landmarks.py`](blender/tesis_landmarks.py) |
| J | Análisis estadístico y boxplots | [`codigo/estadisticas.py`](codigo/estadisticas.py) y [`codigo/boxplots_fstt_por_grupo.py`](codigo/boxplots_fstt_por_grupo.py) |

Los anexos se explican en [docs/anexos.md](docs/anexos.md). El orden de la tabla coincide con los **Anexos A–J de la memoria de título**, no con versiones anteriores del repositorio. Para la carga inicial, consulta la [guía de publicación en GitHub](docs/publicacion.md).

## Decisiones metodológicas fundamentales

- **DICOM:** se verifican metadatos y compatibilidad espacial; se admiten pilas de TC escalares con **al menos 50 cortes utilizables**. No se convierten objetos derivados RGB en intensidades tomográficas aparentes.
- **Segmentación:** máscara corporal inicial en **−500 UH**. El umbral óseo corresponde al **percentil 95 de los vóxeles válidos del volumen (intensidad > −2000 UH)** y después se aplica **dentro de la máscara corporal**. Otsu y *k*-means se registran solo como descriptores; no determinan la etiqueta ósea.
- **Mallas:** `marching_cubes` sobre máscaras binarias al nivel **0,5**, conservando el espaciado físico; **20 iteraciones de suavizado laplaciano** en las superficies finales. La convergencia se examina desde mallas **sin suavizado previo**, no desde mallas ya procesadas.
- **Landmarks:** ubicación anatómica inicial y correcciones por la observadora; apoyo automatizado en la proyección y medición entre superficies.
- **Estadística:** modelos factoriales OLS con inferencia robusta **HC3**, correcciones Benjamini–Hochberg y Holm según el contraste, además de análisis exploratorio de circunferencia cervical utilizada (**CCU**) continua ajustado por edad y sexo. El script conserva otras funciones que **no constituyen resultados principales de la tesis**.

Para más detalles, consultar la [metodología computacional](docs/metodologia.md) y la [correspondencia con el manuscrito](docs/tesis.md).

## Instalación y ejecución

Se requiere un entorno Python compatible con los paquetes de [`requirements.txt`](requirements.txt). El código del complemento se instala **dentro de Blender**, con sus propias dependencias (incluido `openpyxl`). La cabecera del complemento declara **Blender 5.0.0** como versión de referencia, pero no se dispone de una validación independiente de compatibilidad con otras versiones.

Configurar una carpeta de trabajo **fuera del repositorio**, por ejemplo en PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
$env:FSTT_DATA_ROOT = 'D:\datos_autorizados'
python codigo/1_filtrado_estadistica.py --modo todos
```

El procedimiento completo y los comandos de cada etapa aparecen en [docs/ejecucion.md](docs/ejecucion.md). **No ejecute el flujo completo sin datos autorizados y una planilla de control de calidad apropiada.**

## Comprobaciones que se pueden hacer sin datos clínicos

```powershell
python -m unittest discover -s tests -v
python -m compileall -q codigo blender tests
```

Estas verificaciones comprueban configuración y selección de carpetas, además de sintaxis. **No prueban** segmentación anatómica, congruencia geométrica, concordancia interobservador ni predicción forense. Véase [alcance de reproducibilidad](docs/alcance_reproducibilidad.md).

## Privacidad y uso responsable

Los archivos DICOM, NIfTI, PLY, imágenes de control, tablas de medición, resultados individuales, registros con códigos clínicos y cualquier material identificable deben permanecer en un entorno autorizado. El archivo `.gitignore` ayuda a prevenir inclusiones accidentales, pero **no sustituye una revisión manual** antes de cada publicación. Los registros de salida de los scripts pueden contener rutas y códigos de casos; tampoco se deben publicar sin revisión.

## Autoría, licencia y citación

El código se presenta como **desarrollo propio de la autora para esta memoria**, implementado con apoyo de herramientas computacionales y programación asistida. No constituye una implementación ajena atribuida a terceros.

Los metadatos para citar el código se encuentran en [`CITATION.cff`](CITATION.cff). La versión archivada en Zenodo tiene el DOI [`10.5281/zenodo.23550333`](https://doi.org/10.5281/zenodo.23550333). El código se distribuye bajo la [licencia MIT](LICENSE). Esta licencia cubre el código fuente y no autoriza el acceso, uso ni redistribución de datos clínicos, mallas anatómicas, mediciones individuales ni de la tesis.
