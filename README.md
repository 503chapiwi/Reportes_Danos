# Procesador de Reportes de Daños – MAGA / DICORER

Aplicación web (Streamlit) que descarga los reportes de daños enviados desde
KoboCollect y genera la **Planilla de Daños** en el formato oficial de MAGA:
pestañas *Daños agrícolas*, *Daños pecuarios* y *Daños a infraestructura*,
más una pestaña *Contactos* con los agricultores.

## Cómo funciona

1. Los extensionistas llenan el formulario *Reporte de Daños* en KoboCollect
   (funciona sin internet; se envía al recuperar señal).
2. Un usuario de sede departamental o municipal entra a la aplicación, elige
   el período (un mes o un rango de fechas) y el área, y descarga el Excel.
3. La aplicación consulta la API de Kobo en ese momento, así que siempre usa
   los datos más recientes (se guardan en caché 5 minutos; el botón
   *Actualizar datos de Kobo* fuerza una nueva descarga).

Cada usuario solo puede descargar su área: nacional, un departamento o un
municipio, según su configuración.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `app.py` | La interfaz web. |
| `procesador.py` | Descarga de Kobo, transformación y llenado del Excel. |
| `plantilla_danos.xlsx` | La planilla oficial de MAGA, vacía. Se llena una copia en cada descarga. |
| `catalogos.json` | Etiquetas de respaldo (cultivos, fenómenos, etc.). Si Kobo está disponible, se usan las del formulario publicado. |
| `generar_clave.py` | Genera el hash de una contraseña de usuario. |
| `.streamlit/secrets.toml.ejemplo` | Modelo de la configuración (token y usuarios). |

## Configuración

**1. Token de Kobo.** Use una cuenta **institucional de MAGA** (no personal),
para que el sistema siga funcionando cuando cambie el personal. La cuenta debe
ser dueña del formulario o tener permiso de *Ver envíos*. El token está en la
configuración de la cuenta de Kobo, sección *API Key*.

**2. ID del formulario (`asset_uid`).** Está en la URL del formulario en Kobo:
`.../#/forms/`**`vjNPn3ojBDmR6mMJsFEeCp`**`/summary`.

**3. Usuarios.** Para cada sede, ejecute `python generar_clave.py`, escriba la
contraseña y pegue el resultado en `secrets.toml`:

```toml
[usuarios.sacatepequez]
clave = "pbkdf2_sha256$200000$..."
departamento = "3"        # código MAGA del departamento, o "todos"
municipio = "315"         # opcional: limita a un municipio
```

Para quitar el acceso a alguien, borre su bloque.

## Publicar en Streamlit Community Cloud (gratis)

1. Suba esta carpeta a un repositorio **privado** de GitHub (el `.gitignore`
   ya evita que se suba `secrets.toml`).
2. En share.streamlit.io: *Create app* → elija el repositorio → archivo
   principal `app.py`.
3. En *Advanced settings → Secrets*, pegue el contenido de su `secrets.toml`.
4. Comparta la URL con las sedes. Sin usuario y contraseña no se ve ningún dato.

Para probar en una computadora: `pip install -r requirements.txt`, copie
`.streamlit/secrets.toml.ejemplo` como `.streamlit/secrets.toml`, complételo y
ejecute `streamlit run app.py`.

## Mantenimiento

- **Años nuevos:** no requiere cambios. El selector de año crece solo.
- **Nuevo cultivo, fenómeno o especie:** agréguelo en el formulario de Kobo y
  publíquelo. La aplicación toma las etiquetas del formulario publicado.
  Agréguelo también en *DATOS MATRIZ* de `plantilla_danos.xlsx` si quiere que
  aparezca en las listas del Excel.
- **Cambios en la planilla oficial:** si MAGA cambia columnas, hay que ajustar
  `_llenar_agricola` / `_llenar_pec_infra` en `procesador.py`. Las columnas
  están escritas letra por letra para que sea fácil de seguir.
- **Cambios en el formulario:** los campos se leen por su nombre (`name`), no
  por su grupo, así que mover preguntas entre grupos no rompe nada. Si cambia
  el grupo de `fecha_reporte`, `departamento` o `municipio`, actualice
  `RUTA_FECHA`, `RUTA_DEPTO` y `RUTA_MUNI` al inicio de `procesador.py`.
- **Control de calidad:** en Kobo, un supervisor puede marcar un envío como
  *No aprobado* (columna *Validation*). La aplicación lo excluye por defecto.

## Detalles del cálculo

- **Agrícola:** pérdida = precio × quintales perdidos × ponderación por etapa,
  igual que la planilla original (fórmulas en las celdas).
- **Pecuario:** pérdida = cantidad × precio unitario.
- **Infraestructura:** pérdida = valor × % de daño ÷ 100. *La planilla original
  multiplicaba por el porcentaje sin dividir entre 100; confirmar con DICORER.*
- Los códigos de departamento, municipio y uso de suelo se escriben como
  valores (vienen del formulario), en vez de las fórmulas BUSCARV de la
  plantilla, que apuntaban a un archivo externo.
- Fechas en formato día-mes-año.
