# Publicación del repositorio en GitHub

Estas indicaciones corresponden a una **publicación nueva**, sin historial de Git previo. El ZIP entregado contiene código y documentación, no las tomografías ni los resultados individuales.

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

## Después de publicar

1. Copiar la dirección **real** del repositorio desde el navegador y reemplazar en la memoria la marca `[incorporar URL pública verificada]` que aparece en los anexos.
2. Añadir en `CITATION.cff` la línea `repository-code: "https://github.com/USUARIO/fstt-subadultos-chile"` **con la URL real**, no el texto de ejemplo.
3. Verificar nuevamente enlaces de README y documentación.
4. Opcional: crear una versión (release) etiquetada y citar esa versión estable desde la memoria.

**Aclaración:** GitHub público permite **ver** el código. Para otorgar derechos de reutilización y redistribución se necesita definir expresamente una licencia.
