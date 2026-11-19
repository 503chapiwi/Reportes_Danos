"""Streamlit: Reporte de Daños MAGA -> Planilla de Daños (Excel)."""
import calendar
import unicodedata
from datetime import date

import pandas as pd
import streamlit as st

import procesador as P
from generar_clave import verificar

st.set_page_config(page_title='Planilla de Daños – MAGA', page_icon='🌾', layout='centered')

MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto',
         'septiembre', 'octubre', 'noviembre', 'diciembre']


def _sin_tildes(t):
    return unicodedata.normalize('NFKD', str(t)).encode('ascii', 'ignore').decode().replace(' ', '_')


# ----------------------------------------------------------------- config
try:
    KOBO = st.secrets['kobo']
    USUARIOS = st.secrets['usuarios']
except (KeyError, FileNotFoundError):
    st.error('Falta la configuración (secrets). Vea README.md, sección "Configuración".')
    st.stop()


# ----------------------------------------------------------------- login
def pantalla_login():
    st.title('Planilla de Daños')
    st.caption('Ministerio de Agricultura, Ganadería y Alimentación – DICORER')
    with st.form('login'):
        u = st.text_input('Usuario').strip().lower()
        c = st.text_input('Contraseña', type='password')
        if st.form_submit_button('Ingresar', type='primary'):
            datos = USUARIOS.get(u)
            if datos and verificar(c, datos.get('clave', '')):
                st.session_state.usuario = u
                st.rerun()
            st.error('Usuario o contraseña incorrectos.')


if 'usuario' not in st.session_state:
    pantalla_login()
    st.stop()

usuario = st.session_state.usuario
perfil = USUARIOS[usuario]
alc_depto = str(perfil.get('departamento', '')).strip()        # 'todos' or a code
alc_muni = str(perfil.get('municipio', '')).strip() or None    # optional code


@st.cache_data(ttl=3600, show_spinner=False)
def catalogos():
    return P.cargar_catalogos(KOBO['servidor'], KOBO['token'], KOBO['asset_uid'])


@st.cache_data(ttl=300, show_spinner=False)
def envios(desde, hasta, depto, muni):
    return P.descargar_envios(KOBO['servidor'], KOBO['token'], KOBO['asset_uid'], desde, hasta, depto, muni)


cat = catalogos()
deptos = dict(sorted(cat['departamento'].items(), key=lambda kv: kv[1]))
munis = cat['municipio']
depto_de = lambda cod_muni: str(int(cod_muni) // 100)     # 1612 -> 16, 315 -> 3

with st.sidebar:
    st.markdown(f'**Usuario:** {usuario}')
    if alc_depto == 'todos':
        st.caption('Acceso: todos los departamentos')
    else:
        st.caption(f"Acceso: {deptos.get(alc_depto, alc_depto)}"
                   + (f" – {munis.get(alc_muni, alc_muni)}" if alc_muni else ''))
    if st.button('Cerrar sesión'):
        st.session_state.clear()
        st.rerun()
    st.divider()
    if st.button('Actualizar datos de Kobo'):
        envios.clear(); catalogos.clear()
        st.toast('Se volverán a descargar los datos.')

st.title('Planilla de Daños')

# ----------------------------------------------------------------- filtros
c1, c2 = st.columns(2)
modo = c1.radio('Período', ['Un mes', 'Rango de fechas'], horizontal=True)
hoy = date.today()
if modo == 'Un mes':
    anios = list(range(hoy.year, 2025, -1))
    a = c2.selectbox('Año', anios)
    m = st.selectbox('Mes', range(1, 13), index=hoy.month - 1 if a == hoy.year else 0,
                     format_func=lambda i: MESES[i - 1].capitalize())
    desde, hasta = date(a, m, 1), date(a, m, calendar.monthrange(a, m)[1])
    etiqueta_periodo = f'{MESES[m - 1]}_{a}'
else:
    rango = c2.date_input('Del – al', (hoy.replace(day=1), hoy), format='DD/MM/YYYY')
    if not isinstance(rango, tuple) or len(rango) != 2:
        st.info('Seleccione la fecha inicial y la final.')
        st.stop()
    desde, hasta = rango
    etiqueta_periodo = f"{desde:%d-%m-%Y}_al_{hasta:%d-%m-%Y}"

if alc_depto == 'todos':
    opciones = ['todos'] + list(deptos)
    depto = st.selectbox('Departamento', opciones,
                         format_func=lambda k: 'Todos (nacional)' if k == 'todos' else deptos[k])
    depto = None if depto == 'todos' else depto
else:
    depto = alc_depto
    st.selectbox('Departamento', [depto], format_func=lambda k: deptos.get(k, k), disabled=True)

if alc_muni:
    muni = alc_muni
    st.selectbox('Municipio', [muni], format_func=lambda k: munis.get(k, k), disabled=True)
elif depto:
    lista = sorted((k for k in munis if depto_de(k) == depto), key=lambda k: munis[k])
    muni = st.selectbox('Municipio', ['todos'] + lista,
                        format_func=lambda k: 'Todos los municipios' if k == 'todos' else munis[k])
    muni = None if muni == 'todos' else muni
else:
    muni = None

excluir = st.checkbox('Excluir envíos marcados como "No aprobado" en Kobo', value=True)

# ----------------------------------------------------------------- generar
if st.button('Generar planilla', type='primary', width='stretch'):
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

    if res.total_envios == 0:
        st.warning('No hay reportes para este período y área.')
    else:
        lugar = munis.get(muni_r) if muni_r else (deptos.get(depto_r) if depto_r else 'Nacional')
        nombre = f'Planilla_de_Daños_{_sin_tildes(lugar)}_{etiqueta}.xlsx'
        st.download_button('Descargar planilla (Excel)', P.generar_planilla(res), file_name=nombre,
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           type='primary', width='stretch')

    # follow-up lists for supervisors
    if res.sin_danos:
        with st.expander(f'Municipios que reportaron "sin daños" ({len(res.sin_danos)})'):
            st.dataframe(pd.DataFrame([{
                'Departamento': r['depto'], 'Municipio': r['muni'],
                'Extensionista': r['extensionista'], 'Fecha del reporte': f"{r['fecha_reporte']:%d-%m-%Y}",
            } for r in res.sin_danos]), hide_index=True, width='stretch')

    if depto_r and not muni_r:
        con_reporte = {str(r['muni_cod']) for lst in (res.agricola, res.pecuario, res.infraestructura,
                                                     res.sin_danos) for r in lst}
        faltan = sorted(munis[c] for c in munis if depto_de(c) == depto_r and c not in con_reporte)
        if faltan:
            with st.expander(f'Municipios sin ningún reporte en el período ({len(faltan)})'):
                st.write(', '.join(faltan))
