# Alcance de reproducibilidad y estado de validación

La memoria incorpora análisis sobre una cohorte clínica retrospectiva. **Los datos individuales y los archivos de adquisición no están disponibles en el repositorio**. Su consulta o reutilización exige autorizaciones independientes.

## Comprobaciones disponibles en esta publicación

- **Sintaxis:** los archivos `.py` pueden someterse a compilación estática de Python; ello no implica ejecución satisfactoria de los algoritmos ni compatibilidad de Blender.
- **Pruebas unitarias:** tests con nombres y archivos **sintéticos** verifican que las rutas históricas se puedan resolver y que un ZIP no se confunda con una serie clínica. Son pruebas parciales y no usan imágenes médicas.
- **Documentación:** los anexos B–J identifican el script correspondiente y el Anexo A queda únicamente en el manuscrito.

## Lo que no se ha verificado con esta publicación

- Repetición de los espesores numéricos o modelos publicados, puesto que faltan los datos clínicos originales y las planillas privadas.
- Exactitud anatómica de segmentación o de `marching_cubes` frente a un patrón de referencia independiente.
- Equivalencia entre CCU tridimensional y circunferencia medida mediante cinta.
- Concordancia interobservador, rendimiento clínico o identificación facial a partir del cráneo.
- Funcionamiento del complemento en equipos y versiones de Blender diferentes a la configuración de origen.

## Criterios para reutilizar

Registrar el entorno de ejecución, comprobar geometría DICOM y orientaciones, revisar visualmente las máscaras, seleccionar una serie por sujeto, conservar respaldos de PLY antes de la orientación, validar landmarks y evaluar el error en la nueva muestra. **No se recomienda ejecutar a ciegas ni interpretar salidas estadísticas antes de resolver la auditoría de calidad.**

Los logs, rutas, hojas QC, certificados, imágenes, volúmenes, PLY, renders y resultados por individuo nunca deben subirse inadvertidamente al repositorio público.
