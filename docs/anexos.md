# Correspondencia con los Anexos A–J de la memoria

**Memoria:** *Medición del espesor tisular facial en subadultos chilenos mediante tomografía computarizada: Un enfoque semiautomatizado para la aproximación facial forense* (Diéguez-Burgos, 2026).

Este archivo sigue **exactamente la numeración de anexos del manuscrito**. Los programas son desarrollos propios; las fuentes científicas y las bibliotecas se citan en la memoria. Los insumos clínicos originales no están incluidos.

## Anexo A. Aprobaciones éticas y administrativas

**Ubicación:** memoria de título, no repositorio. El certificado y los documentos administrativos fueron incorporados a la memoria conforme a las autorizaciones aplicables. El repositorio no concede acceso a registros hospitalarios.

## Anexo B. Configuración de rutas y organización de directorios

**Archivo:** [`codigo/config_pipeline.py`](../codigo/config_pipeline.py).

Centraliza las rutas de entrada y salida; la raíz de datos se configura mediante `FSTT_DATA_ROOT`. Pueden sobrescribirse las rutas mediante `FSTT_DICOM_DIR`, `FSTT_RESULTADOS_DIR`, `FSTT_SEGMENT_DIR`, `FSTT_MALLAS_DIR` y, para el complemento, `FSTT_EXCEL_ADDON`.

El módulo intenta resolver rutas históricas registradas en planillas a partir de las carpetas de grupo, sexo, caso y adquisición. **No crea la estructura original ni anonimiza los datos**.

## Anexo C. Validación técnica y selección de series DICOM

**Archivos:** [`codigo/1_filtrado_estadistica.py`](../codigo/1_filtrado_estadistica.py), [`codigo/dicom_reader_robusto.py`](../codigo/dicom_reader_robusto.py).

Revisa modalidad, orientación, posición de los cortes, espaciado, matriz y tipo de objeto DICOM. Agrupa imágenes compatibles, ordena los cortes, excluye posiciones duplicadas cuando corresponde y elige una pila de TC escalar compatible con **al menos 50 cortes utilizables**. Registra auditoría y estadísticas de intensidad. No convierte imágenes derivadas RGB en un volumen HU.

**Entrada:** directorios DICOM autorizados. **Salida:** planilla local de casos válidos, errores y auditoría; no se distribuye.

## Anexo D. Segmentación automatizada de las imágenes tomográficas

**Archivo:** [`codigo/2_segmentacion.py`](../codigo/2_segmentacion.py).

Se delimita inicialmente el cuerpo con umbral −500 UH, relleno de huecos, componente principal y cierre 5×5×5. El script calcula el **P95 sobre los vóxeles del volumen con intensidad > −2000 UH**, y utiliza ese valor para clasificar tejido óseo **dentro de la máscara corporal**. El tejido blando y el fondo reciben otras etiquetas; las máscaras se depuran por componentes conectadas y umbrales mínimos del script. Los valores de Otsu y *k*-means (`k=3`) se conservan como descriptores, **no como criterios definitivos de etiquetado**.

**Entrada:** series seleccionadas del paso C. **Salida:** NIfTI y planillas locales de segmentación; requieren revisión visual.

## Anexo E. Generación de representaciones para el control de calidad visual

**Archivo:** [`codigo/3_renders.py`](../codigo/3_renders.py).

Genera vistas estandarizadas de las segmentaciones de tejido blando y tejido óseo. Su finalidad es permitir inspeccionar continuidad, cobertura y posibles artefactos. **La observadora decide** qué serie representa mejor cada paciente y registra esa selección en la planilla de control de calidad. El programa no realiza selección anatómica automática.

**Entrada:** segmentaciones y planilla de seguimiento. **Salida:** renders PNG y registros locales de revisión.

## Anexo F. Extracción y suavizado de mallas tridimensionales

**Archivo:** [`codigo/4_mallasply.py`](../codigo/4_mallasply.py).

Procesa **solo** las reconstrucciones que la observadora marcó como representativas. Extrae las superficies de etiquetas ósea y blanda con `marching_cubes` a nivel 0,5, respetando el espaciado `(z,y,x)` del arreglo y transformándolo en coordenadas físicas. Aplica **20 iteraciones de suavizado laplaciano** y recalcula las normales antes de exportar PLY. **No suaviza los volúmenes ni modifica la segmentación**.

**Entrada:** NIfTI seleccionados y planilla QC. **Salida:** mallas PLY y rutas registradas en la planilla.

## Anexo G. Análisis de convergencia del suavizado laplaciano

**Archivo:** [`codigo/analisisconvergencia.py`](../codigo/analisisconvergencia.py).

Calcula cambios geométricos a partir de **mallas de referencia sin suavizado previo**, recorriendo distintas cantidades de iteraciones de suavizado. Considera distancia de vértices, centroides de caras y orientación de normales, según la implementación. Se empleó para justificar las 20 iteraciones, complementadas con evaluación visual del detalle anatómico.

**Entrada:** mallas de referencia no publicadas. **Salida:** métricas, figuras y eventualmente mallas derivadas, que tampoco se publican al corresponder a un caso individual.

## Anexo H. Preprocesamiento y orientación anatómica de las mallas

**Archivo:** [`codigo/preprocesar_mallas.py`](../codigo/preprocesar_mallas.py).

Aplica de manera consistente a hueso y tejido blando las rotaciones **X +90°** e **Y −90°**. La operación cambia la orientación de las coordenadas, **no la escala**, y debe aplicarse **una sola vez**. El script puede modificar las PLY guardadas; es imprescindible conservar una copia de respaldo.

**Entrada:** PLY y planilla QC. **Salida:** mallas orientadas para su uso posterior en Blender.

## Anexo I. Complemento de Blender para landmarks y medición

**Archivo:** [`blender/tesis_landmarks.py`](../blender/tesis_landmarks.py).

Carga las mallas ósea y blanda de los casos, permite identificar y corregir manualmente landmarks, realiza cálculos geométricos de proyección y distancia, y registra mediciones con estados, sesión, observador y réplica en un archivo Excel local. También incluye la medición exploratoria de **CCU** sobre la superficie cervical. **La identificación anatómica no está automatizada**.

El complemento depende de la API `bpy` de Blender; no se ejecuta como un script Python ordinario. Requiere un Blender compatible y el paquete `openpyxl` instalado en el Python de Blender.

## Anexo J. Procesamiento estadístico y generación de resultados

**Archivos:** [`codigo/estadisticas.py`](../codigo/estadisticas.py), [`codigo/boxplots_fstt_por_grupo.py`](../codigo/boxplots_fstt_por_grupo.py).

Identifica registros válidos, reproduce tablas descriptivas y aplica inferencia robusta HC3 en modelos factoriales de sexo × rango etario, correcciones Benjamini–Hochberg y Holm, evaluación intraobservador, análisis bilateral y asociación exploratoria de la CCU continua ajustada por edad cronológica y sexo.

El programa contiene **análisis categóricos de CCU complementarios** que no forman parte de la interpretación principal de la memoria; requieren validación antes de utilizarlos. Las salidas estadísticas se generan desde el libro de mediciones autorizado y **no acompañan la publicación de código**.

## Reproducibilidad y privacidad

Los scripts pueden inspeccionarse y ejecutarse sobre datos permitidos, pero las salidas originales de la memoria **no se pueden reproducir desde este repositorio por sí solo**. No se incluyen DICOM, NIfTI, PLY, planillas Excel individuales, renders anatómicos, rutas clínicas, logs, contraseñas ni datos identificatorios. La planilla QC, la selección anatómica y el control visual son componentes indispensables del flujo original.

Más información: [ejecución](ejecucion.md), [metodología](metodologia.md) y [limitaciones](alcance_reproducibilidad.md).
