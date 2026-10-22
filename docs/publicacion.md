# Publicación del repositorio en GitHub

Estas indicaciones documentan la preparación y comprobación de la publicación inicial. El repositorio publicado contiene código y documentación, no las tomografías ni los resultados individuales.

## Antes de hacer público el contenido

1. Descomprimir `fstt-subadultos-chile_PUBLICABLE.zip` y **trabajar con la carpeta extraída**; no subir el ZIP como si fuera todo el repositorio.
2. Verificar que la carpeta solo contenga: `codigo/`, `blender/`, `docs/`, `tests/`, `.gitignore`, `README.md`, `CITATION.cff` y `requirements.txt`.
3. Revisar que no se hayan agregado imágenes de pacientes, ficheros DICOM/NIfTI/PLY, Excel, informes privados, logs o datos clínicos a esa carpeta.
4. Decidir si se autorizará reutilización mediante una licencia. Si se quiere permitir uso y modificación con atribución, la autora puede optar por MIT para el código, pero esa decisión requiere su aprobación expresa. **En esta preparación no se añadió ninguna licencia**.

## Publicación con GitHub Desktop (Windows)

1. Abrir [GitHub Desktop](https://desktop.github.com/) e iniciar sesión en GitHub.
2. En **File > Add local repository**, elegir la carpeta `fstt-subadultos-chile` extraída. Si aparece el mensaje de que todavía no es un repositorio Git, elegir **Create a Repository** en esa carpeta.
3. Revisar en la pestaña **Changes** que solo se incluyan los archivos esperados; crear el primer commit (por ejemplo, `Publicación inicial del código de tesis`).
4. Hacer clic en **Publish repository**. Elegir el nombre `fstt-subadultos-chile` y desmarcar **Keep this code private** **solo cuando se hayan verificado** los pasos anteriores.
5. Abrir el repositorio en el navegador y comprobar que las carpetas y el README se ven correctamente.

Guías oficiales de GitHub: [agregar un repositorio local](https://docs.github.com/es/desktop/adding-and-cloning-repositories/adding-a-repository-from-your-local-computer-to-github-desktop?platform=windows), [publicar el proyecto](https://docs.github.com/es/desktop/adding-and-cloning-repositories/adding-an-existing-project-to-github-using-github-desktop).

## Estado de la publicación

El repositorio publicado es [`jodibuu/fstt-subadultos-chile`](https://github.com/jodibuu/fstt-subadultos-chile). Su versión estable inicial es [`v1.0.0`](https://github.com/jodibuu/fstt-subadultos-chile/releases/tag/v1.0.0) y está archivada en Zenodo con el DOI [`10.5281/zenodo.23550333`](https://doi.org/10.5281/zenodo.23550333).

Antes de cada nueva release:

1. Revisar que no se hayan incluido datos clínicos, resultados individuales ni archivos derivados.
2. Verificar los enlaces del README y la documentación.
3. Actualizar la versión y los metadatos de `CITATION.cff` cuando corresponda.
4. Crear una release etiquetada y actualizar el registro de Zenodo si se quiere archivar esa versión.

**Aclaración:** GitHub público permite **ver** el código. Para otorgar derechos de reutilización y redistribución se necesita definir expresamente una licencia.
