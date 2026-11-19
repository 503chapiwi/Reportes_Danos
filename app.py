"""Streamlit: Reporte de Daños MAGA -> Planilla de Daños (Excel) y oficios (Word)."""
import calendar
import io
import unicodedata
from datetime import date, time, timedelta

import pandas as pd
import streamlit as st

import acceso as A
import oficios as O
import procesador as P

st.set_page_config(page_title='Reporte de Daños – MAGA', page_icon='🌾', layout='centered')

MESES = O.MESES
DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
NACIONAL, SEDE = 'nacional', 'sede'


def _sin_tildes(t):
    return unicodedata.normalize('NFKD', str(t)).encode('ascii', 'ignore').decode().replace(' ', '_')


def _seccion(nombre):
    try:
        return {str(k): v for k, v in st.secrets[nombre].items()}
    except (KeyError, FileNotFoundError):
        return {}


# ----------------------------------------------------------------- config
try:
    _secciones = list(st.secrets.keys())
except Exception as e:                      # no secrets at all, or a TOML syntax error
    st.error(f'No se pudieron leer los secrets: {type(e).__name__}: {e}')
    st.stop()
_faltan = [s for s in ('kobo', 'claves') if s not in _secciones]
if _faltan:
    st.error('Falta la configuración (secrets). Secciones que faltan: '
             + ', '.join(f'[{s}]' for s in _faltan)
             + '. Secciones encontradas: ' + (', '.join(f'[{s}]' for s in _secciones) or 'ninguna') + '.')
    st.stop()
KOBO = st.secrets['kobo']
CLAVES = {str(k): str(v) for k, v in st.secrets['claves'].items()}
_faltan = [k for k in ('servidor', 'token', 'asset_uid') if not str(KOBO.get(k, '')).strip()]
if _faltan:
    st.warning('En [kobo] faltan: ' + ', '.join(_faltan) + '. La aplicación no podrá descargar datos.')
SEMILLA = str(_seccion('acceso').get('semilla', ''))
ROTACION = {k: int(v) for k, v in _seccion('rotacion').items()}
JEFES = _seccion('jefes')
OFICIO_CFG = _seccion('oficios')

CAT_LOCAL = P.cargar_catalogos()
depto_de = lambda cod_muni: str(int(cod_muni) // 100)     # 1612 -> 16, 315 -> 3


def munis_de(cat, depto):
    return sorted(((k, v) for k, v in cat['municipio'].items() if depto_de(k) == str(depto)),
                  key=lambda kv: kv[1])


def clave_muni(cod):
    return A.clave_municipal(SEMILLA, cod, ROTACION.get(str(cod), 1))


# ----------------------------------------------------------------- login
def pantalla_login():
    st.title('Reporte de Daños')
    st.caption('Ministerio de Agricultura, Ganadería y Alimentación – DICORER')
    deptos_login = dict(sorted(CAT_LOCAL['departamento'].items(), key=lambda kv: kv[1]))
    opciones = [k for k in deptos_login if k in CLAVES] + ([NACIONAL] if NACIONAL in CLAVES else [])
    d = st.selectbox('Departamento', opciones,
                     format_func=lambda k: 'Nacional (todos)' if k == NACIONAL else deptos_login[k])
    m = SEDE
    if d != NACIONAL and SEMILLA:
        lista = [SEDE] + [k for k, _ in munis_de(CAT_LOCAL, d)]
        m = st.selectbox('Oficina', lista, format_func=lambda k: 'Sede departamental' if k == SEDE
                         else f"AMER {CAT_LOCAL['municipio'][k]}")
    c = st.text_input('Contraseña', type='password')
    if st.button('Ingresar', type='primary'):
        ok = A.igual(c, CLAVES.get(d, '\0'))                      # departmental password
        if not ok and m != SEDE:
            ok = A.coincide(c, clave_muni(m))                      # municipal password
        if ok:
            st.session_state.acceso = (d, None if m == SEDE else m)
            st.rerun()
        st.error('Contraseña incorrecta.')


if 'acceso' not in st.session_state:
    pantalla_login()
    st.stop()

acc_depto, acc_muni = st.session_state.acceso


@st.cache_data(ttl=3600, show_spinner=False)
def catalogos():
    return P.cargar_catalogos(KOBO['servidor'], KOBO['token'], KOBO['asset_uid'])


@st.cache_data(ttl=300, show_spinner=False)
def envios(desde, hasta, depto, muni):
    return P.descargar_envios(KOBO['servidor'], KOBO['token'], KOBO['asset_uid'], desde, hasta, depto, muni)


cat = catalogos()
deptos = dict(sorted(cat['departamento'].items(), key=lambda kv: kv[1]))
munis = cat['municipio']

with st.sidebar:
    if acc_depto == NACIONAL:
        st.markdown('**Acceso:** Nacional')
    else:
        st.markdown(f"**Acceso:** {deptos.get(acc_depto, acc_depto)}")
        st.caption(f"AMER {munis.get(acc_muni, acc_muni)}" if acc_muni else 'Sede departamental')
    if st.button('Cerrar sesión'):
        st.session_state.clear()
        st.rerun()
    st.divider()
    if st.button('Actualizar datos de Kobo'):
        envios.clear(); catalogos.clear()
        st.toast('Se volverán a descargar los datos.')

st.title('Reporte de Daños')

# ----------------------------------------------------------------- filtros
hoy = date.today()
modo = st.radio('Período', ['Semana', 'Mes', 'Rango de fechas'], horizontal=True)
if modo == 'Semana':
    fin = st.date_input('Último día de la semana (fecha del informe)', hoy, format='DD/MM/YYYY')
    desde, hasta = fin - timedelta(days=6), fin
    st.caption(f'Semana del {O.semana(desde, hasta)}')
    etiqueta_periodo = f'{desde:%d-%m}_al_{hasta:%d-%m-%Y}'
elif modo == 'Mes':
    c1, c2 = st.columns(2)
    a = c1.selectbox('Año', list(range(hoy.year, 2025, -1)))
    m = c2.selectbox('Mes', range(1, 13), index=hoy.month - 1 if a == hoy.year else 0,
                     format_func=lambda i: MESES[i - 1].capitalize())
    desde, hasta = date(a, m, 1), date(a, m, calendar.monthrange(a, m)[1])
    etiqueta_periodo = f'{MESES[m - 1]}_{a}'
else:
    rango = st.date_input('Del – al', (hoy.replace(day=1), hoy), format='DD/MM/YYYY')
    if not isinstance(rango, tuple) or len(rango) != 2:
        st.info('Seleccione la fecha inicial y la final.')
        st.stop()
    desde, hasta = rango
    etiqueta_periodo = f'{desde:%d-%m-%Y}_al_{hasta:%d-%m-%Y}'

if acc_depto == NACIONAL:
    depto = st.selectbox('Departamento', ['todos'] + list(deptos),
                         format_func=lambda k: 'Todos (nacional)' if k == 'todos' else deptos[k])
    depto = None if depto == 'todos' else depto
else:
    depto = acc_depto
    st.selectbox('Departamento', [depto], format_func=lambda k: deptos.get(k, k), disabled=True)

if acc_muni:
    muni = acc_muni
    st.selectbox('Municipio', [muni], format_func=lambda k: munis.get(k, k), disabled=True)
elif depto:
    muni = st.selectbox('Municipio', ['todos'] + [k for k, _ in munis_de(cat, depto)],
                        format_func=lambda k: 'Todos los municipios' if k == 'todos' else munis[k])
    muni = None if muni == 'todos' else muni
else:
    muni = None

excluir = st.checkbox('Excluir envíos marcados como "No aprobado" en Kobo', value=True)

# ----------------------------------------------------------------- generar
if st.button('Generar reporte', type='primary', width='stretch'):
    try:
        with st.spinner('Descargando datos de Kobo…'):
            datos = envios(desde, hasta, depto, muni)
    except P.ErrorKobo as e:
        st.error(str(e))
        st.stop()
    res = P.procesar(datos, cat, desde, hasta, depto, muni, excluir_no_aprobados=excluir)
    st.session_state.resultado = (res, desde, hasta, depto, muni, etiqueta_periodo)

if 'resultado' in st.session_state:
    res, desde_r, hasta_r, depto_r, muni_r, etiqueta = st.session_state.resultado
    st.subheader(f"{desde_r:%d-%m-%Y} al {hasta_r:%d-%m-%Y}")

    k = st.columns(4)
    k[0].metric('Reportes', res.total_envios)
    k[1].metric('Filas agrícolas', len(res.agricola))
    k[2].metric('Filas pecuarias', len(res.pecuario))
    k[3].metric('Infraestructura', len(res.infraestructura))
    st.caption(f'{len(res.contactos)} contactos de agricultores'
               + (f' · {res.excluidos_no_aprobados} envíos excluidos por "No aprobado"'
                  if res.excluidos_no_aprobados else ''))

    lugar = munis.get(muni_r) if muni_r else (deptos.get(depto_r) if depto_r else 'Nacional')

    # ---- planilla (Excel)
    st.markdown('#### Matriz Oficial de Daños (Excel)')
    if res.total_envios == 0:
        st.warning('No hay reportes para este período y área.')
    else:
        st.download_button('Descargar planilla (Excel)', P.generar_planilla(res),
                           file_name=f'Planilla_de_Daños_{_sin_tildes(lugar)}_{etiqueta}.xlsx',
                           mime=XLSX, width='stretch')

    # ---- oficios (Word)
    if depto_r:
        st.markdown('#### Oficio para firma (Word)')
        jefe_def = JEFES.get(str(depto_r), '')
        c1, c2 = st.columns([1, 2])
        jefe_titulo = c1.text_input('Título del Jefe Departamental', 'Ingeniero Agrónomo')
        jefe_nombre = c2.text_input('Nombre del Jefe Departamental', jefe_def)
        c1, c2, c3 = st.columns(3)
        fecha_oficio = c1.date_input('Fecha del oficio', hasta_r, format='DD/MM/YYYY')
        folios = c2.number_input('Folios adjuntos', 1, 999, 2)
        hora = c3.time_input('Hora del informe', time(8, 30)) if muni_r else None
        datos_of = dict(fecha_oficio=fecha_oficio, jefe_titulo=jefe_titulo, jefe_nombre=jefe_nombre,
                        folios=folios, hora=f'{hora:%H:%M}' if hora else '')
        nombre_depto = deptos.get(depto_r, '')

        if muni_r:
            if not jefe_nombre:
                st.info('Escriba el nombre del Jefe Departamental para generar el oficio.')
            else:
                doc = O.oficio_municipal(res, muni_r, munis[muni_r], depto_r, nombre_depto,
                                         desde_r, hasta_r, datos_of)
                cod = f'{O.codigo_municipio_corto(muni_r)}-{int(depto_r):02d}'
                st.download_button('Descargar oficio municipal (Word)', doc, mime=DOCX, type='primary',
                                   width='stretch',
                                   file_name=f'{cod}_Oficio_daños_{_sin_tildes(munis[muni_r])}_{fecha_oficio:%d-%m-%Y}.docx')
        else:
            c1, c2 = st.columns([1, 2])
            sub_titulo = c1.text_input('Título del Subdirector', OFICIO_CFG.get('sub_titulo', 'Licenciado'))
            sub_nombre = c2.text_input('Nombre del Subdirector', OFICIO_CFG.get('sub_nombre', ''))
            c1, c2 = st.columns(2)
            referencia = c1.text_input('Ref. (opcional)')
            oficio_no = c2.text_input('Oficio No. (opcional)', placeholder='467-2026-SD-MAGA-SAC')
            datos_of.update(sub_titulo=sub_titulo, sub_nombre=sub_nombre, referencia=referencia,
                            oficio_no=oficio_no)
            if not (jefe_nombre and sub_nombre):
                st.info('Escriba el nombre del Jefe Departamental y del Subdirector para generar los oficios.')
            else:
                lista = munis_de(cat, depto_r)
                st.download_button('Descargar oficio departamental (Word)',
                                   O.oficio_departamental(res, lista, depto_r, nombre_depto,
                                                          desde_r, hasta_r, datos_of),
                                   file_name=f'Oficio_SD-{int(depto_r):02d}_daños_{_sin_tildes(nombre_depto)}_{fecha_oficio:%d-%m-%Y}.docx',
                                   mime=DOCX, type='primary', width='stretch')
                st.download_button(f'Descargar los {len(lista)} oficios municipales (ZIP)',
                                   O.oficios_municipales_zip(res, lista, depto_r, nombre_depto,
                                                             desde_r, hasta_r, dict(datos_of, folios=2)),
                                   file_name=f'Oficios_municipales_{_sin_tildes(nombre_depto)}_{fecha_oficio:%d-%m-%Y}.zip',
                                   mime='application/zip', width='stretch')

    # ---- follow-up lists for supervisors
    if res.sin_danos:
        with st.expander(f'Municipios que reportaron "sin daños" ({len(res.sin_danos)})'):
            st.dataframe(pd.DataFrame([{
                'Departamento': r['depto'], 'Municipio': r['muni'],
                'Extensionista': r['extensionista'], 'Fecha del reporte': f"{r['fecha_reporte']:%d-%m-%Y}",
            } for r in res.sin_danos]), hide_index=True, width='stretch')

    if depto_r and not muni_r:
        con_reporte = {str(r['muni_cod']) for lst in (res.agricola, res.pecuario, res.infraestructura,
                                                     res.sin_danos) for r in lst}
        faltan = [n for c, n in munis_de(cat, depto_r) if c not in con_reporte]
        if faltan:
            with st.expander(f'Municipios sin ningún reporte en el período ({len(faltan)})'):
                st.write(', '.join(faltan))

# ----------------------------------------------------------------- claves AMER (solo sede/nacional)
if SEMILLA and not acc_muni:
    lista_d = list(deptos) if acc_depto == NACIONAL else [acc_depto]
    filas = [{'Departamento': deptos[d], 'Municipio': n, 'Código': c, 'Contraseña': clave_muni(c)}
             for d in lista_d for c, n in munis_de(cat, d)]
    with st.expander(f'Contraseñas de las AMER ({len(filas)})'):
        st.caption('Cada AMER ingresa eligiendo su municipio y escribiendo su contraseña. '
                   'Comparta a cada oficina solo la suya.')
        df = pd.DataFrame(filas)
        st.dataframe(df, hide_index=True, width='stretch')
        buf = io.StringIO(); df.to_csv(buf, index=False)
        st.download_button('Descargar lista (CSV)', buf.getvalue().encode('utf-8-sig'),
                           file_name='contraseñas_AMER.csv', mime='text/csv')
