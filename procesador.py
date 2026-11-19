"""Procesador de Reportes de Daños (Kobo -> planilla MAGA).

Pure logic, no Streamlit: fetch submissions from the KoboToolbox API,
turn them into planilla rows, and fill MAGA's own template.
"""
from __future__ import annotations

import io
import json
from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import openpyxl
import requests
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

AQUI = Path(__file__).parent
PLANTILLA = AQUI / 'plantilla_danos.xlsx'
CATALOGOS = AQUI / 'catalogos.json'

# Full paths in the Kobo JSON for fields used in the server-side query.
# Only needed for the query; everything else is matched by the last path
# segment, so renaming a group in the form doesn't break the processor.
RUTA_FECHA = 'datos_generales/fecha_reporte'
RUTA_DEPTO = 'datos_generales/departamento'
RUTA_MUNI = 'datos_generales/municipio'

FMT_FECHA = 'dd-mm-yyyy'


# ----------------------------------------------------------------- Kobo API
class ErrorKobo(Exception):
    pass


def _get(url, token, params=None):
    try:
        r = requests.get(url, headers={'Authorization': f'Token {token}'}, params=params, timeout=60)
    except requests.RequestException as e:
        raise ErrorKobo(f'No se pudo conectar con Kobo: {e}') from e
    if r.status_code in (401, 403):
        raise ErrorKobo('Kobo rechazó el token (401/403). Revise el token y que la cuenta tenga acceso al formulario.')
    if r.status_code == 404:
        raise ErrorKobo('Kobo no encontró el formulario (404). Revise el asset_uid y la URL del servidor.')
    return r


def descargar_envios(servidor, token, asset_uid, desde: date, hasta: date,
                     depto: str | None = None, muni: str | None = None, tam_pagina=5000):
    """All submissions in the date range (and departamento/municipio if given)."""
    url = f"{servidor.rstrip('/')}/api/v2/assets/{asset_uid}/data/"
    rango = {RUTA_FECHA: {'$gte': desde.isoformat(), '$lte': hasta.isoformat()}}
    q = {'$and': [{'$or': [rango, {RUTA_FECHA: {'$exists': False}}]}]}
    if depto:
        q['$and'].append({RUTA_DEPTO: str(depto)})
    if muni:
        q['$and'].append({RUTA_MUNI: str(muni)})

    params = {'format': 'json', 'limit': tam_pagina, 'query': json.dumps(q)}
    r = _get(url, token, params)
    if r.status_code == 400:            # server didn't accept the query -> filter locally
        params.pop('query')
        r = _get(url, token, params)
    if r.status_code != 200:
        raise ErrorKobo(f'Error de Kobo ({r.status_code}): {r.text[:300]}')

    envios = []
    while True:
        datos = r.json()
        envios.extend(datos.get('results', []))
        siguiente = datos.get('next')
        if not siguiente:
            break
        r = _get(siguiente, token)
        if r.status_code != 200:
            raise ErrorKobo(f'Error de Kobo al paginar ({r.status_code}).')
    return envios


def cargar_catalogos(servidor=None, token=None, asset_uid=None):
    """Choice labels: bundled copy, updated from the live form when reachable."""
    cat = json.loads(CATALOGOS.read_text(encoding='utf-8'))
    if servidor and token and asset_uid:
        try:
            r = _get(f"{servidor.rstrip('/')}/api/v2/assets/{asset_uid}/", token, {'format': 'json'})
            if r.status_code == 200:
                for ch in r.json().get('content', {}).get('choices', []):
                    lab = ch.get('label')
                    lab = lab[0] if isinstance(lab, list) and lab else lab
                    if ch.get('list_name') and ch.get('name') is not None and lab:
                        cat.setdefault(ch['list_name'], {})[str(ch['name'])] = lab
        except ErrorKobo:
            pass
    return cat


# ----------------------------------------------------------------- transform
def _plano(d: dict) -> dict:
    return {k.split('/')[-1]: v for k, v in d.items()}


def _num(v):
    if v in (None, ''):
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return int(f) if f.is_integer() else f


def _fecha(v):
    if not v:
        return None
    try:
        return datetime.strptime(str(v)[:10], '%Y-%m-%d').date()
    except ValueError:
        return None


@dataclass
class Resultado:
    agricola: list = field(default_factory=list)
    pecuario: list = field(default_factory=list)
    infraestructura: list = field(default_factory=list)
    contactos: list = field(default_factory=list)
    sin_danos: list = field(default_factory=list)
    excluidos_no_aprobados: int = 0
    total_envios: int = 0


def procesar(envios, cat, desde: date, hasta: date, depto=None, muni=None,
             excluir_no_aprobados=True) -> Resultado:
    lab = lambda lista, v: cat.get(lista, {}).get(str(v), v) if v not in (None, '') else None
    res = Resultado()

    for e in envios:
        s = _plano(e)

        # filters re-applied locally (server query may be unavailable)
        f_rep = _fecha(s.get('fecha_reporte')) or _fecha(s.get('fecha_mon_del')) \
            or _fecha(s.get('_submission_time'))
        if not f_rep or not (desde <= f_rep <= hasta):
            continue
        if depto and str(s.get('departamento')) != str(depto):
            continue
        if muni and str(s.get('municipio')) != str(muni):
            continue
        estado = (s.get('_validation_status') or {}).get('uid', '')
        if excluir_no_aprobados and estado == 'validation_status_not_approved':
            res.excluidos_no_aprobados += 1
            continue
        res.total_envios += 1

        base = dict(
            depto_cod=_num(s.get('departamento')),
            depto=s.get('departamento_nombre') or lab('departamento', s.get('departamento')),
            muni_cod=_num(s.get('municipio')),
            muni=s.get('municipio_nombre') or lab('municipio', s.get('municipio')),
            extensionista=s.get('nombre_extensionista'),
            fecha_reporte=f_rep,
        )

        if s.get('hubo_danos') == 'no':
            res.sin_danos.append(base)
            continue

        base.update(
            lugar=s.get('lugar_poblado'),
            lat=_num(s.get('latitud')), lon=_num(s.get('longitud')),
            fenomeno=lab('fenomeno', s.get('fenomeno')),
            nombre_fenomeno=s.get('nombre_fenomeno'),
            mon_del=_fecha(s.get('fecha_mon_del')), mon_al=_fecha(s.get('fecha_mon_al')),
        )
        tipo = s.get('tipo_dano')
        cultivos_envio = {}

        if tipo == 'agricola':
            for rep in s.get('cultivo_detalle') or []:
                c = _plano(rep)
                nombre = c.get('cultivo_nombre') or lab('cultivo', c.get('cultivo'))
                cultivos_envio[c.get('cultivo')] = nombre
                res.agricola.append(dict(
                    base, cultivo=nombre, codigo_uso=_num(c.get('codigo_uso')),
                    detalle=c.get('detalle'),
                    aut=_num(c.get('autoconsumo_a')), inter=_num(c.get('intermedia_a')),
                    cons=_num(c.get('consolidada_a')),
                    siembra=_fecha(c.get('fecha_siembra')),
                    etapa=lab('etapa', c.get('etapa_fenologica')), ponderacion=_num(c.get('ponderacion')),
                    ha_afect=_num(c.get('hectareas_afectadas')),
                    produccion=_num(c.get('produccion_quintales')),
                    superficie=_num(c.get('superficie_sembrada')),
                    precio=_num(c.get('precio_quintal')),
                ))
        elif tipo == 'pecuario':
            res.pecuario.append(dict(
                base, especie=lab('especie', s.get('especie')),
                animales=_num(s.get('animales_perdidos')), precio=_num(s.get('precio_unitario_p')),
                aut=_num(s.get('autoconsumo_p')), inter=_num(s.get('intermedia_p')),
                cons=_num(s.get('consolidada_p')),
            ))
        elif tipo == 'infraestructura':
            t = s.get('tipo_infraestructura')
            nombre = s.get('infraestructura_otro') if t == 'otro' else lab('infraestructura', t)
            res.infraestructura.append(dict(
                base, infra=nombre, porcentaje=_num(s.get('porcentaje_dano')),
                valor=_num(s.get('precio_unitario_i')),
                aut=_num(s.get('autoconsumo_i')), inter=_num(s.get('intermedia_i')),
                cons=_num(s.get('consolidada_i')),
            ))

        tipo_lab = lab('tipo_dano', tipo)
        for rep in s.get('agricultores') or []:
            c = _plano(rep)
            detalle_tipo = tipo_lab
            if tipo == 'agricola':
                sel = (c.get('cultivos_contacto') or '').split() or list(cultivos_envio)
                nombres = [cultivos_envio.get(x) or lab('cultivo', x) for x in sel]
                if nombres:
                    detalle_tipo = f"{tipo_lab} ({', '.join(nombres)})"
            elif tipo == 'pecuario' and res.pecuario:
                detalle_tipo = f"{tipo_lab} ({res.pecuario[-1]['especie']})"
            elif tipo == 'infraestructura' and res.infraestructura:
                detalle_tipo = f"{tipo_lab} ({res.infraestructura[-1]['infra']})"
            res.contactos.append(dict(
                nombre=c.get('nombre_agricultor'),
                telefono=c.get('telefono_agricultor') or 'Sin teléfono',
                depto=base['depto'], muni=base['muni'], lugar=base['lugar'],
                tipo=tipo, tipo_detalle=detalle_tipo,
            ))

    orden = lambda r: (str(r.get('depto') or ''), str(r.get('muni') or ''), str(r.get('lugar') or ''),
                       r.get('mon_del') or date.min)
    for lst in (res.agricola, res.pecuario, res.infraestructura, res.contactos):
        lst.sort(key=orden)
    res.sin_danos.sort(key=lambda r: (str(r['depto']), str(r['muni']), r['fecha_reporte']))
    return res


# ----------------------------------------------------------------- Excel
def _copiar_estilo(ws, fila_origen, fila_destino, col_ini, col_fin):
    for c in range(col_ini, col_fin + 1):
        o, d = ws.cell(fila_origen, c), ws.cell(fila_destino, c)
        if o.has_style:
            d._style = copy(o._style)
    h = ws.row_dimensions[fila_origen].height
    if h:
        ws.row_dimensions[fila_destino].height = h


def _insertar_filas(ws, en_fila, n, fila_estilo, col_fin):
    """insert_rows that also moves merged cells and row heights below it."""
    if n <= 0:
        return
    abajo = [mr for mr in ws.merged_cells.ranges if mr.min_row >= en_fila]
    for mr in abajo:
        ws.unmerge_cells(str(mr))
    alturas = {r: d.height for r, d in ws.row_dimensions.items() if r >= en_fila and d.height}
    ws.insert_rows(en_fila, n)
    for r in sorted(alturas, reverse=True):
        ws.row_dimensions[r + n].height = alturas[r]
        ws.row_dimensions[r].height = None
    for mr in abajo:
        mr.shift(0, n)
        ws.merge_cells(str(mr))
    for r in range(en_fila, en_fila + n):
        _copiar_estilo(ws, fila_estilo, r, 1, col_fin)


def _fechas(ws, fila, cols):
    for col in cols:
        ws[f'{col}{fila}'].number_format = FMT_FECHA


def _llenar_agricola(ws, filas):
    PRIMERA, ULTIMA_PLANTILLA = 10, 34
    # the template's rows are formatted unevenly; give every data row row 10's format
    for r in range(PRIMERA + 1, max(ULTIMA_PLANTILLA, PRIMERA + len(filas) - 1) + 1):
        _copiar_estilo(ws, PRIMERA, r, 1, 29)
    for i, d in enumerate(filas):
        r = PRIMERA + i
        vals = {
            'A': i + 1, 'B': d['cultivo'], 'C': d['codigo_uso'], 'D': d['detalle'],
            'E': d['fenomeno'], 'F': d['nombre_fenomeno'], 'G': d['depto_cod'], 'H': d['depto'],
            'I': d['muni_cod'], 'J': d['muni'], 'K': d['lugar'], 'L': d['lat'], 'M': d['lon'],
            'N': d['aut'], 'O': d['inter'], 'P': d['cons'], 'Q': f'=SUM(N{r}:P{r})',
            'R': d['siembra'], 'S': d['etapa'], 'T': d['ponderacion'], 'U': d['ha_afect'],
            'V': d['mon_del'], 'W': d['mon_al'], 'X': d['produccion'], 'Y': d['superficie'],
            'Z': f'=+IFERROR(X{r}/Y{r},0)', 'AA': f'=Z{r}*U{r}', 'AB': d['precio'],
            'AC': f'=ROUND(AB{r}*AA{r}*T{r},2)',
        }
        for col, v in vals.items():
            ws[f'{col}{r}'] = v
        _fechas(ws, r, 'RVW')
    for r in range(PRIMERA + len(filas), ULTIMA_PLANTILLA + 1):
        ws[f'A{r}'] = r - PRIMERA + 1


def _llenar_pec_infra(ws, filas, tipo, contactos):
    PRIMERA, ULTIMA_PLANTILLA, FILA_TOTAL = 8, 57, 58
    extra = max(0, len(filas) - (ULTIMA_PLANTILLA - PRIMERA + 1))
    _insertar_filas(ws, FILA_TOTAL, extra, PRIMERA, 19)
    ultima = ULTIMA_PLANTILLA + extra
    for r in range(PRIMERA + 1, ultima + 1):
        _copiar_estilo(ws, PRIMERA, r, 1, 19)
    total = FILA_TOTAL + extra
    for col in 'KLMNOS':
        ws[f'{col}{total}'] = f'=SUM({col}{PRIMERA}:{col}{ultima})'

    for i, d in enumerate(filas):
        r = PRIMERA + i
        comunes = {
            'A': i + 1, 'C': d['fenomeno'], 'D': d['depto_cod'], 'E': d['depto'],
            'F': d['muni_cod'], 'G': d['muni'], 'H': d['lugar'], 'I': d['lat'], 'J': d['lon'],
            'L': d['aut'], 'M': d['inter'], 'N': d['cons'], 'O': f'=SUM(L{r}:N{r})',
            'P': d['mon_del'], 'Q': d['mon_al'],
        }
        if tipo == 'pecuario':
            comunes.update(B=d['especie'], K=d['animales'], R=d['precio'], S=f'=+R{r}*K{r}')
        else:  # valor x % daño / 100 (same formula as the Kobo form)
            comunes.update(B=d['infra'], K=d['porcentaje'], R=d['valor'], S=f'=ROUND(R{r}*K{r}/100,2)')
        for col, v in comunes.items():
            ws[f'{col}{r}'] = v
        _fechas(ws, r, 'PQ')
    for r in range(PRIMERA + len(filas), ultima + 1):   # numbering for empty rows, as in the original
        ws[f'A{r}'] = r - PRIMERA + 1

    # "Informantes" table at the bottom of the sheet
    inf_ini = 68 + extra
    inf_fin = {'pecuario': 200, 'infraestructura': 73}[tipo] + extra
    for i, c in enumerate(contactos):
        r = inf_ini + i
        if r > inf_fin:
            _copiar_estilo(ws, inf_ini, r, 1, 3)
        ws[f'A{r}'] = i + 1
        ws[f'B{r}'] = c['nombre']
        tel = c['telefono']
        ws[f'C{r}'] = int(tel) if str(tel).isdigit() else tel


def _hoja_contactos(wb, contactos):
    ws = wb.create_sheet('Contactos')
    enc = ['No.', 'Nombre del agricultor', 'Teléfono', 'Municipio', 'Lugar poblado', 'Tipo de daño']
    ws.append(enc)
    for i, c in enumerate(contactos, 1):
        tel = c['telefono']
        ws.append([i, c['nombre'], str(tel), c['muni'], c['lugar'], c['tipo_detalle']])
    borde = Border(*(Side(style='thin', color='999999'),) * 4)
    for celda in ws[1]:
        celda.font = Font(name='Arial', bold=True, color='FFFFFF')
        celda.fill = PatternFill('solid', fgColor='1F4E78')
        celda.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        celda.border = borde
    for fila in ws.iter_rows(min_row=2):
        for celda in fila:
            celda.font = Font(name='Arial', size=10)
            celda.border = borde
    for col, ancho in zip('ABCDEF', [6, 32, 13, 28, 28, 34]):
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = 'A2'
    if contactos:
        ws.auto_filter.ref = f'A1:F{len(contactos) + 1}'
    for fila in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        fila[0].alignment = Alignment(horizontal='center')


def generar_planilla(res: Resultado, fecha_informe: date | None = None) -> bytes:
    wb = openpyxl.load_workbook(PLANTILLA)
    hoy = fecha_informe or date.today()

    agr = wb['Daños agrícolas']
    _llenar_agricola(agr, res.agricola)
    agr['C7'] = hoy; agr['C7'].number_format = FMT_FECHA

    for hoja, tipo, filas in (('Daños pecuarios', 'pecuario', res.pecuario),
                              ('Daños a infraestructura', 'infraestructura', res.infraestructura)):
        ws = wb[hoja]
        _llenar_pec_infra(ws, filas, tipo, [c for c in res.contactos if c['tipo'] == tipo])
        ws['D5'] = hoy; ws['D5'].number_format = FMT_FECHA

    _hoja_contactos(wb, res.contactos)
    wb.active = wb.sheetnames.index('Daños agrícolas')
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
