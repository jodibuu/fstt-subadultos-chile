# Ejecución orientativa y requisitos

Estas instrucciones describen el orden de ejecución del código, **no un ejemplo completamente reproducible sin datos**. El procesamiento integral requiere imágenes de TC autorizadas, planillas de seguimiento y revisión anatómica de cada caso.

## Entorno Python y configuración

El código no establece un entorno de versiones exactas. El paquete original se desarrolló con herramientas Python compatibles con las versiones de la memoria; al reutilizarlo se deben registrar y verificar las versiones realmente instaladas. La cabecera del complemento indica **Blender 5.0.0**. No se ha probado aquí Blender ni la ejecución integral con DICOM.

En Windows PowerShell, desde la raíz del repositorio:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
$env:FSTT_DATA_ROOT = 'D:\datos_autorizados'
```

La ruta `D:\datos_autorizados` es un **ejemplo genérico**, no una ruta de la cohorte real. La carpeta configurada debe ser externa al repositorio público. Los archivos se organizan, cuando corresponde, como:

```text
<FSTT_DATA_ROOT>/
├── 1_DATOS_BRUTOS/<grupo>/<sexo>/<caso>/<serie>/
├── 2_PROCESAMIENTO/
├── 3_MALLAS_PLY/
└── 5_RESULTADOS/
```

Las rutas específicas pueden sobrescribirse mediante las variables `FSTT_DICOM_DIR`, `FSTT_SEGMENT_DIR`, `FSTT_MALLAS_DIR`, `FSTT_RESULTADOS_DIR` y `FSTT_EXCEL_ADDON`.

## Secuencia y comandos

**B/C — Validar series DICOM** (genera planilla de validación y auditoría):

```powershell
python codigo/1_filtrado_estadistica.py --modo todos
```

**D — Segmentar** (revisar disponibles y procesar faltantes):

```powershell
python codigo/2_segmentacion.py --modo verificar
python codigo/2_segmentacion.py --modo incompletos
```

**E — Generar renders para control visual**:

```powershell
python codigo/3_renders.py
```

**Intervención obligatoria:** revisar las imágenes de control, elegir **una sola serie representativa por paciente** y registrar la decisión en la planilla QC. Las columnas esperadas para la selección incluyen `Paciente`, `Carpeta`, `Ruta salida segmentación` y `¿Es la mejor carpeta del paciente? (1 = Sí)`; se puede registrar `¿Descartar? (1 = Sí)`. La existencia de una segmentación no implica aceptación anatómica.

**F — Generar mallas de series seleccionadas**:

```powershell
python codigo/4_mallasply.py --excel 'D:\datos_autorizados\control_calidad.xlsx'
```

**G — Análisis de convergencia, si se desea repetir la evaluación**: establecer `FSTT_MALLA_REFERENCIA_HUESO`, `FSTT_MALLA_REFERENCIA_BLANDO` y, si corresponde, `FSTT_CONVERGENCIA_SALIDA` con rutas a **mallas sin suavizado previo**. No aplicar esta comprobación sobre las PLY que ya pasaron por el suavizado rutinario.

**H — Orientar mallas** (**modifica archivos PLY**, respaldar antes):

```powershell
python codigo/preprocesar_mallas.py --excel 'D:\datos_autorizados\control_calidad.xlsx'
```

**I — Medición en Blender**: instalar o abrir `blender/tesis_landmarks.py` en una versión de Blender compatible. Instalar `openpyxl` en el entorno Python interno de Blender si falta. Configurar `FSTT_EXCEL_ADDON` para el libro de mediciones autorizado o seleccionar la ruta mediante las opciones de la herramienta. **Comprobar escala, orientación, landmarks y correspondencia de las superficies antes de guardar resultados.**

**J — Estadísticas** (usar libro exportado, no suministrado aquí):

```powershell
python codigo/estadisticas_tesis_HC3.py 'D:\datos_autorizados\mediciones.xlsx' --salida 'D:\datos_autorizados\ANALISIS_HC3_publicacion_local'
```

El comando anterior es solo un ejemplo de rutas. Por defecto, el script crea una **subcarpeta con fecha** en el directorio de la entrada y guarda allí `RESULTADOS_HC3.xlsx`; si se indica `--salida`, esa carpeta debe estar vacía.

Para los boxplots, indicar la ruta del Excel resultante (evita copiarlo a una ubicación distinta):

```powershell
$env:FSTT_BOXPLOTS_EXCEL = 'D:\datos_autorizados\ANALISIS_HC3_publicacion_local\RESULTADOS_HC3.xlsx'
$env:FSTT_BOXPLOTS_DIR = 'D:\datos_autorizados\figuras_boxplots'
python codigo/boxplots_fstt_por_grupo.py
```

Las variables `FSTT_BOXPLOTS_EXCEL` y `FSTT_BOXPLOTS_DIR` son ajustes de portabilidad: **no cambian las transformaciones estadísticas ni la construcción de las figuras**.

## Comprobaciones sin tomografías

```powershell
python -m unittest discover -s tests -v
python -m compileall -q codigo blender tests
```

Las pruebas con datos sintéticos son limitadas: comprueban resolución de rutas y selección de carpetas, **no la exactitud de espesores ni la ejecución de Blender**.

**Advertencia de privacidad:** Excel, renders, DICOM, mallas, logs y salidas clínicas deben mantenerse fuera de GitHub. Las rutas y los identificadores pueden aparecer en reportes locales, aunque el código fuente no contenga datos personales.
