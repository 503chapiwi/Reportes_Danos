"""Oficios de daños (Word): municipal (AMER -> Sede) y departamental (Sede -> Subdirección)."""
from __future__ import annotations

import io
import zipfile
from collections import OrderedDict
from datetime import date
from pathlib import Path

from docxtpl import DocxTemplate

AQUI = Path(__file__).parent
PLANTILLA_MUNI = AQUI / 'plantilla_oficio_municipal.docx'
PLANTILLA_DEPTO = AQUI / 'plantilla_oficio_departamental.docx'

MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto',
         'septiembre', 'octubre', 'noviembre', 'diciembre']


# ----------------------------------------------------------------- formato
def fecha_larga(d: date) -> str:
    return f'{d.day:02d} de {MESES[d.month - 1]} de {d.year}'


def semana(desde: date, hasta: date) -> str:
    if desde.year != hasta.year:
        return (f'{desde.day:02d} de {MESES[desde.month - 1]} del {desde.year} al '
                f'{hasta.day:02d} de {MESES[hasta.month - 1]} del {hasta.year}')
    if desde.month != hasta.month:
        return f'{desde.day:02d} de {MESES[desde.month - 1]} al {hasta.day:02d} de {MESES[hasta.month - 1]} del {hasta.year}'
    return f'{desde.day:02d} al {hasta.day:02d} de {MESES[hasta.month - 1]} del {hasta.year}'


def _entero(x):
    return f'{int(round(x or 0)):,}'


def _dec(x):
    x = round(x or 0, 2)
    return f'{x:,.2f}'.rstrip('0').rstrip('.') if x else '0'


def _q(x):
    return f'Q{x:,.2f}' if x else '0'


def _lista(valores):
    """unique, in order; 'Maíz criollo, maíz blanco y frijol' style."""
    vistos = list(OrderedDict.fromkeys(v for v in valores if v))
    if not vistos:
        return 'Ninguno'
    vistos = [vistos[0]] + [v[0].lower() + v[1:] for v in vistos[1:]]
    return ', '.join(vistos)


def _texto_lista(valores):
    l = _lista(valores)
    partes = l.split(', ')
    return l if len(partes) < 2 else ', '.join(partes[:-1]) + ' y ' + partes[-1]


def _fam(r):
    return (r.get('aut') or 0) + (r.get('inter') or 0) + (r.get('cons') or 0)


def _perdida_agr(r):
    try:
        return round((r['precio'] or 0) * (r['produccion'] / r['superficie']) * (r['ha_afect'] or 0)
                     * (r['ponderacion'] or 0), 2)
    except (TypeError, ZeroDivisionError):
        return 0.0


def _perdida_pec(r):
    return round((r['animales'] or 0) * (r['precio'] or 0), 2)


def _perdida_inf(r):
    return round((r['valor'] or 0) * (r['porcentaje'] or 0) / 100, 2)


def codigo_municipio_corto(cod_muni) -> str:
    return f'{int(cod_muni) % 100:02d}'


# ----------------------------------------------------------------- tablas
def _filas_agr(res, munis_cod_nombre):
    filas, tot = [], dict(f=0, ha=0.0, q=0.0, fen=[], cult=[])
    for cod, nombre in munis_cod_nombre:
        rs = [r for r in res.agricola if str(r['muni_cod']) == str(cod)]
        f = sum(_fam(r) for r in rs); ha = sum(r['ha_afect'] or 0 for r in rs)
        q = sum(_perdida_agr(r) for r in rs)
        fen = [r['fenomeno'] for r in rs]; cult = [r['cultivo'] for r in rs]
        filas.append(dict(municipio=nombre, familias=_entero(f), hectareas=_dec(ha), perdida=_q(q),
                          fenomeno=_lista(fen), cultivo=_lista(cult)))
        tot['f'] += f; tot['ha'] += ha; tot['q'] += q; tot['fen'] += fen; tot['cult'] += cult
    total = dict(familias=_entero(tot['f']), hectareas=_dec(tot['ha']), perdida=_q(tot['q']),
                 fenomeno=_lista(tot['fen']), cultivo=_lista(tot['cult']))
    return filas, total, tot


def _filas_pec(res, cods):
    grupos = OrderedDict()
    for r in res.pecuario:
        if str(r['muni_cod']) in cods:
            grupos.setdefault((r['muni'], r['especie']), []).append(r)
    filas, tf, ta, tq = [], 0, 0, 0.0
    for (muni, especie), rs in grupos.items():
        f = sum(_fam(r) for r in rs); a = sum(r['animales'] or 0 for r in rs)
        q = sum(_perdida_pec(r) for r in rs)
        filas.append(dict(municipio=muni, familias=_entero(f), animales=_entero(a), perdida=_q(q),
                          fenomeno=_lista(r['fenomeno'] for r in rs), especie=especie or ''))
        tf += f; ta += a; tq += q
    return filas, dict(familias=_entero(tf), animales=_entero(ta), perdida=_q(tq)), (tf, ta, tq, grupos)


def _filas_inf(res, cods):
    filas, tf, tq = [], 0, 0.0
    for r in res.infraestructura:
        if str(r['muni_cod']) not in cods:
            continue
        f = _fam(r); q = _perdida_inf(r)
        filas.append(dict(municipio=r['muni'], familias=_entero(f), porcentaje=f"{_dec(r['porcentaje'])} %",
                          perdida=_q(q), fenomeno=r['fenomeno'] or 'Ninguno', tipo=r['infra'] or ''))
        tf += f; tq += q
    return filas, dict(familias=_entero(tf), perdida=_q(tq)), (tf, tq)


def _extensionista(res, cod):
    rs = [r for lst in (res.agricola, res.pecuario, res.infraestructura, res.sin_danos)
          for r in lst if str(r['muni_cod']) == str(cod) and r.get('extensionista')]
    return max(rs, key=lambda r: r['fecha_reporte'])['extensionista'] if rs else ''


# ----------------------------------------------------------------- render
def _render(plantilla, ctx) -> bytes:
    doc = DocxTemplate(plantilla)
    doc.render(ctx, autoescape=True)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def oficio_municipal(res, cod_muni, nombre_muni, cod_depto, nombre_depto, desde, hasta, datos) -> bytes:
    """datos: fecha_oficio, jefe_titulo, jefe_nombre, hora, folios"""
    cods = {str(cod_muni)}
    agr, _, tot = _filas_agr(res, [(cod_muni, nombre_muni)])
    pec, _, (pf, pa, pq, grupos) = _filas_pec(res, cods)
    inf, _, (inf_f, inf_q) = _filas_inf(res, cods)

    if tot['ha'] or tot['f']:
        texto_agr = (f"sí se reportaron daños agrícolas en {_dec(tot['ha'])} "
                     f"{'hectárea' if tot['ha'] == 1 else 'hectáreas'} de "
                     f"{_texto_lista(tot['cult']).lower()}, afectando a {_entero(tot['f'])} familias, "
                     f"con una pérdida estimada de {_q(tot['q'])}.")
    else:
        texto_agr = 'no se reportaron daños agrícolas.'
    if pec:
        animales = ', '.join(f"{_entero(sum(r['animales'] or 0 for r in rs))} {esp.lower()}"
                             for (_, esp), rs in grupos.items())
        texto_pec = (f'Se reportó la pérdida de {animales}, afectando a {_entero(pf)} familias, '
                     f'con una pérdida estimada de {_q(pq)}.')
    else:
        texto_pec = 'No hay reporte de daños pecuarios.'
    if inf:
        tipos = ', '.join(f"{x['tipo'].lower()} ({x['porcentaje']})" for x in inf)
        texto_inf = (f'Se reportaron daños en {tipos}, afectando a {_entero(inf_f)} familias, '
                     f'con una pérdida estimada de {_q(inf_q)}.')
    else:
        texto_inf = 'No hay reporte de daños de infraestructura agropecuaria.'

    f_of = datos['fecha_oficio']
    ctx = dict(
        municipio=nombre_muni, municipio_mayus=nombre_muni.upper(), departamento=nombre_depto,
        fecha_larga=fecha_larga(f_of), semana=semana(desde, hasta), hora=datos.get('hora', ''),
        codigo_oficio=f"AMER-{codigo_municipio_corto(cod_muni)}-{int(cod_depto):02d}-DAÑOS-{f_of:%d-%m-%Y}",
        jefe_titulo=datos.get('jefe_titulo', ''), jefe_nombre=datos.get('jefe_nombre', ''),
        folios=f"{int(datos.get('folios') or 2):02d}",
        extensionista=_extensionista(res, cod_muni),
        agr=agr, pec=pec, inf=inf,
        adjunta_matriz=bool(tot['f'] or pec or inf),
        texto_agricola=texto_agr, texto_pecuario=texto_pec, texto_infra=texto_inf,
    )
    return _render(PLANTILLA_MUNI, ctx)


def oficio_departamental(res, munis_cod_nombre, cod_depto, nombre_depto, desde, hasta, datos) -> bytes:
    """munis_cod_nombre: every municipio of the departamento (zeros included).
    datos: fecha_oficio, jefe_titulo, jefe_nombre, sub_titulo, sub_nombre, referencia, oficio_no, folios"""
    cods = {str(c) for c, _ in munis_cod_nombre}
    agr, tot_agr, tot = _filas_agr(res, munis_cod_nombre)
    pec, tot_pec, _ = _filas_pec(res, cods)
    inf, tot_inf, _ = _filas_inf(res, cods)
    f_of = datos['fecha_oficio']
    ctx = dict(
        departamento=nombre_depto, departamento_mayus=nombre_depto.upper(),
        fecha_larga=fecha_larga(f_of), semana=semana(desde, hasta),
        codigo_oficio=f'SD-{int(cod_depto):02d}-DAÑOS-{f_of:%d-%m-%Y}',
        referencia=datos.get('referencia', ''), oficio_no=datos.get('oficio_no', ''),
        sub_titulo=datos.get('sub_titulo', ''), sub_nombre=datos.get('sub_nombre', ''),
        jefe_titulo=datos.get('jefe_titulo', ''), jefe_nombre=datos.get('jefe_nombre', ''),
        folios=f"{int(datos.get('folios') or 1):02d}",
        agr=agr, tot_agr=tot_agr, pec=pec, tot_pec=tot_pec, inf=inf, tot_inf=tot_inf,
        adjunta_matriz=bool(tot['f'] or pec or inf),
    )
    return _render(PLANTILLA_DEPTO, ctx)


def oficios_municipales_zip(res, munis_cod_nombre, cod_depto, nombre_depto, desde, hasta, datos) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for cod, nombre in munis_cod_nombre:
            doc = oficio_municipal(res, cod, nombre, cod_depto, nombre_depto, desde, hasta, datos)
            z.writestr(f'{codigo_municipio_corto(cod)}-{int(cod_depto):02d}_Oficio_daños_{nombre}.docx', doc)
    return buf.getvalue()
