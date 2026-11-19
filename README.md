# Procesador de Reportes de Daños – MAGA / DICORER

Aplicación web (Streamlit) que descarga los reportes de daños enviados desde
KoboCollect y genera:

- La **Planilla de Daños** (Excel) en el formato oficial de MAGA: pestañas
  *Daños agrícolas*, *Daños pecuarios* y *Daños a infraestructura*, más una
  pestaña *Contactos* con los agricultores.
- El **oficio municipal** (Word, para firma): oficio de la AMER a la Sede más
  el *Informe de daños a cultivos, pecuarios y de infraestructura productiva*.
- El **oficio departamental** (Word, para firma): oficio de la Sede a la
  Subdirección, con una fila por cada municipio del departamento (incluso los
  que no tienen daños) y su total. La sede también puede descargar los oficios
  municipales de todas sus AMER en un ZIP.

Las tablas de daños pecuarios e infraestructura solo aparecen en los oficios
cuando hay ese tipo de daño.

## Cómo funciona

1. Los extensionistas llenan el formulario *Reporte de Daños* en KoboCollect
   (funciona sin internet; se envía al recuperar señal).
2. Una sede departamental o una AMER entra a la aplicación, elige el período
   (semana, mes o rango de fechas) y descarga el Excel y el oficio.
3. La aplicación consulta la API de Kobo en ese momento, así que siempre usa
   los datos más recientes (se guardan en caché 5 minutos; el botón
   *Actualizar datos de Kobo* fuerza una nueva descarga).

Cada sede solo puede descargar los datos de su departamento, y cada AMER los
de su municipio.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `app.py` | La interfaz web. |
| `procesador.py` | Descarga de Kobo, transformación y llenado del Excel. |
| `oficios.py` | Llena los oficios municipal y departamental. |
| `acceso.py` | Cálculo de las contraseñas municipales. |
| `plantilla_oficio_municipal.docx`, `plantilla_oficio_departamental.docx` | Plantillas Word de los oficios. Se pueden editar en Word (sin tocar lo que está entre `{{ }}` o `{% %}`). |
| `plantilla_danos.xlsx` | La planilla oficial de MAGA, vacía. Se llena una copia en cada descarga. |
| `catalogos.json` | Etiquetas de respaldo (cultivos, fenómenos, etc.). Si Kobo está disponible, se usan las del formulario publicado. |
| `.streamlit/secrets.toml.ejemplo` | Modelo de la configuración (token y contraseñas). |

## Configuración

**1. Token de Kobo.** Use una cuenta **institucional de MAGA** (no personal),
para que el sistema siga funcionando cuando cambie el personal. La cuenta debe
ser dueña del formulario o tener permiso de *Ver envíos*. El token está en la
configuración de la cuenta de Kobo, sección *API Key*.

**2. ID del formulario (`asset_uid`).** Está en la URL del formulario en Kobo:
`.../#/forms/`**`vjNPn3ojBDmR6mMJsFEeCp`**`/summary`.

**3. Contraseñas.** Hay dos niveles:

- **Sede departamental:** una contraseña por departamento, escrita en la
  sección `[claves]` con el código MAGA (`"3" = "..."`). `nacional` da acceso
  a los 22 departamentos. Solo aparecen al ingresar los departamentos que
  tienen contraseña, así que se pueden ir agregando por etapas.
- **AMER (municipal):** no se escriben. Se calculan a partir de una sola
  `semilla` secreta (sección `[acceso]`), de modo que cada municipio tiene su
  propia contraseña (por ejemplo `K7QM-4TXA`) sin cargar 340 claves. Cada sede
  ve las de sus municipios en la aplicación (sección *Contraseñas de las AMER*)
  y puede descargarlas en CSV para repartirlas. Una AMER solo ve los datos de
  su municipio; la contraseña departamental también entra a cualquier
  municipio del departamento.
- **Cambiar la clave de un municipio:** agregue su código en `[rotacion]` con
  el número 2 (luego 3, etc.). Cambiar la `semilla` cambia todas a la vez.

**4. Nombres para los oficios (opcional).** En `[jefes]` (nombre del Jefe de
cada sede) y `[oficios]` (Subdirector). Aparecen ya escritos en la aplicación
y se pueden corregir antes de descargar.

## Publicar en Streamlit Community Cloud (gratis)

1. Suba esta carpeta a un repositorio **privado** de GitHub (el `.gitignore`
   ya evita que se suba `secrets.toml`).
2. En share.streamlit.io: *Create app* → elija el repositorio → archivo
   principal `app.py`.
3. En *Advanced settings → Secrets*, pegue el contenido de su `secrets.toml`.
4. Comparta la URL con las sedes. Sin contraseña no se ve ningún dato.

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
- **Código del oficio:** `AMER-15-03-DAÑOS-08-10-2026` = número del municipio
  dentro del departamento (315 → 15), código del departamento y fecha del
  oficio. El departamental es `SD-03-DAÑOS-08-10-2026`. *Ref.* y *Oficio No.*
  vienen del registro de correspondencia y se escriben a mano en la aplicación.
- **Familias en los oficios:** suma de las familias de cada fila de la
  planilla. Si un mismo grupo de familias se reportó en dos cultivos, se cuenta
  dos veces (igual que al sumar la planilla).
