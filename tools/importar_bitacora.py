"""
Convierte la bitácora vieja en Excel (bloques por mes y sede) al formato de la
hoja Registros de la app.

Uso:
    python3 tools/importar_bitacora.py Bitacora.xlsx salida/ [--anio 2026] [--correcciones correcciones.json]

  correcciones.json (opcional, NO va al repositorio porque tiene nombres) corrige filas del
  Excel sin editarlo. Cada entrada busca por fecha (dd/mm) y el inicio del nombre del paciente:
    [{"fecha": "06/11", "paciente": "Bruno Ramírez", "estado": "Confirmada", "obs": "Datáfono"},
     {"fecha": "26/09", "paciente": "María Paula Morales", "obs_agregar": "Datáfono"}]
  Campos: estado, cobro, sana, obs (reemplaza) u obs_agregar (se suma a Observaciones).

Genera en la carpeta de salida:
  - importar.xlsx   una pestaña "Importar" con las columnas de Registros. En la hoja de
                    la app: Archivo → Importar → subir → "Insertar hojas nuevas", y luego
                    menú Hábitos → Importar registros.
  - revision.xlsx   pestañas Revisar (filas no importadas y avisos) y Cuadre (totales
                    por mes y sede contra el Excel original).

Reglas:
  - DS = Escazú. Bloques "<Mes> DS | Domicilio | Curridabat".
  - Solo se importan filas con estado "Confirmada"; el resto va a "Revisar".
  - Nuevo / Inicial → Primera vez · Inicial Pareja → Pareja primera vez (dos filas
    seguidas del mismo día se unen en un registro de pareja) · Medición → Mediciones.
  - Datáfono → Tarjeta · SINPE → SINPE · "Descuento 10% (VOLK)" → descuento VOLK; otros
    "Descuento ..." → descuento "Otro" por la diferencia con la tarifa de lista.
  - Columna SANA SANA → seguro Sana Sana (estado del cobro: Pendiente).
  - Nombres incompletos se completan con la hoja "Base de datos" cuando hay un único candidato.
"""
import argparse
import datetime as dt
import hashlib
import json
import os
import re
import sys
import unicodedata

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto',
         'septiembre', 'octubre', 'noviembre', 'diciembre']
DIAS = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo']
SEDES = {'ds': 'Escazú', 'escazu': 'Escazú', 'domicilio': 'Domicilio', 'curridabat': 'Curridabat'}

# Tarifas vigentes (por persona); deben coincidir con TARIFAS_INICIALES de Code.gs.
TARIFAS = {
    'Primera vez': {'Escazú': 50000, 'Curridabat': 40000, 'Domicilio': 35000},
    'Control': {'Escazú': 35000, 'Curridabat': 30000, 'Domicilio': 30000},
    'Pareja primera vez': {'Escazú': 47000, 'Curridabat': 35000, 'Domicilio': 30000},
    'Pareja control': {'Escazú': 32000, 'Curridabat': 27000, 'Domicilio': 25000},
    'Mediciones': {'Escazú': 14000, 'Domicilio': 12000},
}
VOLK = 0.10

VERDE = PatternFill('solid', fgColor='304323')
CREMA = 'F7F2E8'
AMARILLO = PatternFill('solid', fgColor='FBF0C6')


def norm(s):
    s = unicodedata.normalize('NFD', str(s or '')).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', s).strip().lower()


def encabezados_registros():
    """Lee los encabezados de COLUMNAS en src/Code.gs para no desincronizarse."""
    code = open(os.path.join(RAIZ, 'src', 'Code.gs'), encoding='utf-8').read()
    bloque = code[code.index('const COLUMNAS = ['):]
    bloque = bloque[:bloque.index('];')]
    return re.findall(r"key: '(\w+)', header: '([^']+)'", bloque)


def tipo_de(texto):
    t = norm(texto)
    pareja = 'pareja' in t
    if 'medic' in t:
        return 'Mediciones'
    if 'control' in t:
        return 'Pareja control' if pareja else 'Control'
    if 'inicial' in t or 'nuevo' in t or 'primera' in t:
        return 'Pareja primera vez' if pareja else 'Primera vez'
    return None


def forma_de(obs):
    o = norm(obs)
    if 'datafono' in o or 'tarjeta' in o:
        return 'Tarjeta'
    if 'sinpe' in o:
        return 'SINPE'
    return ''


def leer_bloques(ws, avisos):
    """Devuelve filas crudas con mes, sede y columnas por nombre de encabezado."""
    filas = []
    grupos = {}   # columna inicial → {'mes', 'sede', 'cols': {nombre: idx}}
    for fila in ws.iter_rows():
        valores = [c.value for c in fila]
        nuevos = {}
        for i, v in enumerate(valores):
            m = re.match(r'^\s*(\w+)\s+(DS|Domicilio|Curridabat|Escaz[uú])\s*$', str(v or ''), re.I)
            if m and norm(m.group(1)) in MESES:
                nuevos[i] = {'mes': MESES.index(norm(m.group(1))) + 1, 'sede': SEDES[norm(m.group(2))], 'cols': None}
        if nuevos:
            # Los bloques de una misma fila son del mismo mes: si uno dice otro mes, es un error de tipeo.
            meses = [g['mes'] for g in nuevos.values()]
            mayoria = max(set(meses), key=lambda x: (meses.count(x), -meses.index(x)))
            for i, g in nuevos.items():
                if g['mes'] != mayoria:
                    avisos.append((fila[i].coordinate, str(valores[i]),
                                   'El encabezado dice %s pero los demás bloques de la fila son de %s; se usó %s.'
                                   % (MESES[g['mes'] - 1], MESES[mayoria - 1], MESES[mayoria - 1])))
                    g['mes'] = mayoria
            grupos.update(nuevos)
        for inicio, g in grupos.items():
            if norm(valores[inicio] if inicio < len(valores) else '') == 'fecha':
                cols = {}
                j = inicio
                while j < len(valores) and valores[j] not in (None, ''):
                    cols[norm(valores[j])] = j
                    j += 1
                cols['_extra'] = j   # columna sin encabezado justo después (p. ej. "Ya se pagó el 40%")
                g['cols'] = cols
                continue
            if not g['cols']:
                continue
            c = g['cols']
            get = lambda k: valores[c[k]] if k in c and c[k] < len(valores) else None
            paciente = get('paciente')
            fecha_txt = get('fecha')
            if not paciente or not fecha_txt or norm(fecha_txt).startswith('total'):
                continue
            extra = valores[c['_extra']] if c['_extra'] < len(valores) else None
            filas.append({
                'celda': fila[inicio].coordinate, 'mes': g['mes'], 'sede': g['sede'],
                'fecha_txt': str(fecha_txt).strip(), 'paciente': str(paciente).strip(),
                'tipo_txt': str(get('tipo de servicio') or '').strip(),
                'estado': str(get('estado') or '').strip(), 'cobro': get('cobro'),
                'sana': get('sana sana'), 'obs': str(get('observaciones') or '').strip(),
                'extra': str(extra or '').strip(),
            })
    return filas


def nombres_base(wb):
    if 'Base de datos' not in wb.sheetnames:
        return []
    ws = wb['Base de datos']
    cab = [norm(c.value) for c in ws[1]]
    if 'paciente' not in cab:
        return []
    i = cab.index('paciente')
    return [str(r[i]).strip() for r in ws.iter_rows(min_row=2, values_only=True) if r[i]]


def completar_nombre(nombre, base):
    """Usa el nombre completo de la Base de datos si hay exactamente un candidato."""
    n = norm(nombre)
    if any(norm(b) == n for b in base):
        return nombre, ''   # ya está completo; se respeta cómo viene escrito en la agenda
    cand = [b for b in base if norm(b).startswith(n + ' ')]
    if len(cand) == 1:
        return cand[0], 'Nombre completado con Base de datos (era "%s")' % nombre
    return nombre, ''


def aplicar_correcciones(crudas, correcciones, anio, avisos):
    """Corrige filas crudas del Excel según el archivo de correcciones (ver docstring del módulo)."""
    for c in correcciones:
        dia, mes = [int(x) for x in c['fecha'].split('/')[:2]]
        nombre = norm(c['paciente'])
        hallados = [f for f in crudas if f['mes'] == mes and re.search(r'(\d{1,2})', f['fecha_txt'])
                    and int(re.search(r'(\d{1,2})', f['fecha_txt']).group(1)) == dia
                    and norm(f['paciente']).startswith(nombre)]
        if len(hallados) != 1:
            avisos.append(('', c['paciente'], 'Corrección del %s: %d filas coinciden; no se aplicó.' % (c['fecha'], len(hallados))))
            continue
        f = hallados[0]
        for k in ('estado', 'cobro', 'sana', 'obs'):
            if k in c:
                f[k] = c[k]
        if c.get('obs_agregar'):
            f['obs'] = (f['obs'] + ', ' if f['obs'] else '') + c['obs_agregar']
        f['corregida'] = True


def convertir(ruta, anio, hoy, correcciones=None):
    wb = openpyxl.load_workbook(ruta, data_only=True)
    ws = wb['Agenda 2026'] if 'Agenda 2026' in wb.sheetnames else wb.worksheets[0]
    base = nombres_base(wb)
    registros, revisar, avisos = [], [], []
    crudas = leer_bloques(ws, avisos)
    aplicar_correcciones(crudas, correcciones or [], anio, avisos)
    pendientes_pareja = {}
    for f in crudas:
        motivo = []
        dia = re.search(r'(\d{1,2})', f['fecha_txt'])
        fecha = None
        if dia:
            try:
                fecha = dt.date(anio, f['mes'], int(dia.group(1)))
            except ValueError:
                pass
        if not fecha:
            motivo.append('fecha no reconocida')
        else:
            dia_txt = norm(f['fecha_txt'].split()[0])
            if dia_txt in DIAS and DIAS.index(dia_txt) != fecha.weekday():
                avisos.append((f['celda'], f['paciente'], '"%s" no coincide con el día de la semana del %s; se usó el número.'
                               % (f['fecha_txt'], fecha.strftime('%d/%m/%Y'))))
        tipo = tipo_de(f['tipo_txt'])
        if not tipo:
            motivo.append('tipo "%s" no reconocido' % f['tipo_txt'])
        estado = norm(f['estado'])
        if estado != 'confirmada':
            motivo.append('estado "%s"' % f['estado'].strip())
        if motivo:
            revisar.append(dict(f, fecha=fecha, motivo='; '.join(motivo)))
            continue

        nombre, nota_nombre = completar_nombre(f['paciente'], base)
        cobro = int(f['cobro'] or 0)
        sana = int(f['sana'] or 0)
        obs = f['obs']
        volk = 'volk' in norm(obs) and 'descuento' in norm(obs)
        forma = forma_de(obs)
        notas = ['Importado de la bitácora en Excel']
        if obs:
            notas.append('Obs.: ' + obs)
        if f['extra']:
            notas.append(f['extra'])
        if f.get('corregida'):
            notas.append('Corregido al importar')
        if nota_nombre:
            notas.append(nota_nombre)
        if cobro and not forma:
            avisos.append((f['celda'], nombre, 'Sin forma de pago en Observaciones; quedó vacía.'))
        if 'regal' in norm(obs):
            avisos.append((f['celda'], nombre, 'Dice "Regalía" pero el Cobro es %d; se importó con ese monto.' % cobro))
        en_obs = re.search(r'\((\d+)\s*mil\)', norm(obs))
        if en_obs and int(en_obs.group(1)) * 1000 != cobro:
            avisos.append((f['celda'], nombre, 'Observaciones dice %s mil pero el Cobro es %d; se usó el Cobro.'
                           % (en_obs.group(1), cobro)))

        reg = {
            'fecha': fecha, 'hora': '', 'paciente': nombre, 'paciente2': '', 'tipo': tipo, 'ubicacion': f['sede'],
            'conSeguro': 'Sí' if sana else 'No', 'aseguradora': 'Sana Sana' if sana else '',
            'pago1': cobro, 'forma1': forma if cobro else '', 'pago2': '', 'forma2': '',
            'montoAseguradora': sana or '', 'estadoAseguradora': 'Pendiente' if sana else '',
            'descuento': 'VOLK' if volk else ('Otro' if 'descuento' in norm(obs) else ''),
            'notas': notas, 'celda': f['celda'],
        }

        if tipo.startswith('Pareja'):
            # Se une con la fila anterior de la misma sede y tipo si las fechas difieren a lo sumo un día.
            clave = (f['sede'], tipo)
            previa = pendientes_pareja.get(clave)
            if previa and abs((previa['fecha'] - fecha).days) > 1:
                avisos.append((previa['celda'], previa['paciente'], 'Cita de pareja sin segunda persona confirmada; revisar.'))
                pendientes_pareja.pop(clave)
                previa = None
            if previa:
                p = pendientes_pareja.pop(clave)
                if p['fecha'] != fecha:
                    avisos.append((reg['celda'], reg['paciente'], 'Pareja con fechas distintas (%s y %s); se usó %s.'
                                   % (p['fecha'].strftime('%d/%m'), fecha.strftime('%d/%m'), p['fecha'].strftime('%d/%m'))))
                p['paciente2'] = reg['paciente']
                p['pago2'] = reg['pago1']
                p['forma2'] = reg['forma1']
                p['notas'] += [n for n in reg['notas'] if n not in p['notas']]
                p['celda'] += ' + ' + reg['celda']
            else:
                pendientes_pareja[clave] = reg
                registros.append(reg)
            continue
        registros.append(reg)

    for reg in pendientes_pareja.values():
        avisos.append((reg['celda'], reg['paciente'], 'Cita de pareja sin segunda persona confirmada; revisar.'))

    # Tarifas de lista y descuento, igual que la app
    for r in registros:
        personas = 2 if r['tipo'].startswith('Pareja') else 1
        base_t = TARIFAS.get(r['tipo'], {}).get(r['ubicacion'])
        r['tarifa'] = base_t * personas if base_t else ''
        if r['descuento'] == 'VOLK' and r['tarifa']:
            r['montoDescuento'] = round(r['tarifa'] * VOLK)
        elif r['descuento'] == 'Otro' and r['tarifa']:
            # Descuento puntual: lo que faltó para llegar a la tarifa de lista
            r['montoDescuento'] = max(0, r['tarifa'] - (r['pago1'] or 0) - (r['pago2'] or 0) - (r['montoAseguradora'] or 0)) or ''
        else:
            r['montoDescuento'] = ''
        r['totalPaciente'] = (r['pago1'] or 0) + (r['pago2'] or 0)
        r['totalFacturado'] = r['totalPaciente'] + (r['montoAseguradora'] or 0)
    registros.sort(key=lambda r: (r['fecha'], r['ubicacion']))
    ahora = dt.datetime.now().replace(microsecond=0)
    for r in registros:
        # ID estable: el mismo registro tiene el mismo ID aunque el Excel crezca, así reimportar no duplica.
        clave = '|'.join([r['fecha'].isoformat(), r['ubicacion'], norm(r['paciente']), r['tipo']])
        r['id'] = 'H' + hashlib.sha1(clave.encode()).hexdigest()[:7]
        r['creado'] = r['modificado'] = ahora
        r['eventoId'] = ''
        r['notas'] = ' · '.join(r['notas'])
    return registros, revisar, avisos, crudas


def cuadre(registros, crudas):
    """Compara por mes y sede lo importado contra la suma de la columna Cobro (+ Sana Sana) del Excel."""
    tot = {}
    for f in crudas:
        k = (f['mes'], f['sede'])
        t = tot.setdefault(k, {'excel': 0, 'import': 0, 'n_excel': 0, 'n_import': 0})
        t['excel'] += int(f['cobro'] or 0) + int(f['sana'] or 0)
        t['n_excel'] += 1
    for r in registros:
        t = tot.setdefault((r['fecha'].month, r['ubicacion']), {'excel': 0, 'import': 0, 'n_excel': 0, 'n_import': 0})
        t['import'] += r['totalFacturado']
        t['n_import'] += 1
    return sorted(tot.items())


def escribir(registros, revisar, avisos, filas_cuadre, salida):
    os.makedirs(salida, exist_ok=True)
    cols = encabezados_registros()
    wb = openpyxl.Workbook()

    def cabecera(ws, titulos, anchos):
        ws.append(titulos)
        for i, c in enumerate(ws[1], 1):
            c.font = Font(bold=True, color=CREMA)
            c.fill = VERDE
            c.alignment = Alignment(vertical='center')
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = anchos[i - 1] if i <= len(anchos) else 14
        ws.freeze_panes = 'A2'

    ws = wb.active
    ws.title = 'Importar'
    cabecera(ws, [h for _, h in cols], [8, 11, 7, 30, 24, 18, 12, 10, 12, 13, 10, 13, 10, 14, 10, 14, 13, 15, 15, 14, 60, 14, 17, 17])
    for r in registros:
        ws.append([r.get(k, '') for k, _ in cols])
        fila = ws.max_row
        for j, (k, _) in enumerate(cols, 1):
            c = ws.cell(fila, j)
            if k == 'fecha':
                c.number_format = 'dd/mm/yyyy'
            elif k in ('creado', 'modificado'):
                c.number_format = 'dd/mm/yyyy hh:mm'
            elif k in ('tarifa', 'montoDescuento', 'pago1', 'pago2', 'totalPaciente', 'montoAseguradora', 'totalFacturado'):
                c.number_format = '"₡"#,##0'

    ruta_imp = os.path.join(salida, 'importar.xlsx')
    wb.save(ruta_imp)

    wb = openpyxl.Workbook()
    wr = wb.active
    wr.title = 'Revisar'
    cabecera(wr, ['Celda en Excel', 'Fecha', 'Sede', 'Paciente', 'Tipo', 'Estado', 'Cobro', 'Sana Sana', 'Observaciones', 'Motivo'],
             [10, 11, 12, 30, 16, 15, 10, 10, 34, 30])
    for f in revisar:
        wr.append([f['celda'], f['fecha'], f['sede'], f['paciente'], f['tipo_txt'], f['estado'], f['cobro'], f['sana'],
                   ' · '.join(x for x in (f['obs'], f['extra']) if x), f['motivo']])
        wr.cell(wr.max_row, 2).number_format = 'dd/mm/yyyy'
    wr.append([])
    wr.append(['Avisos de filas importadas'])
    wr.cell(wr.max_row, 1).font = Font(bold=True)
    for celda, nombre, texto in avisos:
        wr.append([celda, '', '', nombre, '', '', '', '', '', texto])
        for c in wr[wr.max_row]:
            c.fill = AMARILLO

    wc = wb.create_sheet('Cuadre')
    cabecera(wc, ['Mes', 'Sede', 'Filas en Excel', 'Suma Excel (cobro + Sana Sana)', 'Registros importados', 'Total importado', 'Diferencia'],
             [12, 12, 13, 28, 20, 16, 14])
    for (mes, sede), t in filas_cuadre:
        wc.append([MESES[mes - 1].capitalize(), sede, t['n_excel'], t['excel'], t['n_import'], t['import'], t['excel'] - t['import']])
        for j in (4, 6, 7):
            wc.cell(wc.max_row, j).number_format = '"₡"#,##0'
    wc.append([])
    wc.append(['La diferencia corresponde a filas que no están "Confirmada" (canceladas, por reprogramar o pendientes); ver pestaña Revisar.'])

    ruta_rev = os.path.join(salida, 'revision.xlsx')
    wb.save(ruta_rev)
    return ruta_imp, ruta_rev


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('excel')
    ap.add_argument('salida')
    ap.add_argument('--anio', type=int, default=2026)
    ap.add_argument('--correcciones', help='JSON con correcciones a filas del Excel')
    a = ap.parse_args()
    correcciones = json.load(open(a.correcciones, encoding='utf-8')) if a.correcciones else []
    registros, revisar, avisos, crudas = convertir(a.excel, a.anio, dt.date.today(), correcciones)
    rx, rc = escribir(registros, revisar, avisos, cuadre(registros, crudas), a.salida)
    print('Importados: %d registros · A revisar: %d filas · Avisos: %d' % (len(registros), len(revisar), len(avisos)))
    print(rx)
    print(rc)


if __name__ == '__main__':
    sys.exit(main())
