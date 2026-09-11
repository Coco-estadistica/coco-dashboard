#!/usr/bin/env python3
"""
Análisis de cosechas de la originación de crédito de Districolmotos.

Lee el Excel de cosechas, homologa los cortes mensuales, reconstruye el triángulo
de cosechas en meses en libros (MOB), exporta las tablas fuente y actualiza los
datos embebidos en Cosechas_Districolmotos.html.

Uso:
    python3 analizar_cosechas.py [ruta/al/Cosechas_2026.xlsx]

Reglas que respeta (no las cambies sin pensarlo dos veces):
  - Un dato faltante es s/d, nunca cero. Las series llevan None, no 0.
  - Las cosechas solo se comparan a MOB constante.
  - El numerador es el CAPITAL VENCIDO (SALDO EN MORA), no el saldo del crédito.
  - El denominador es la colocación original de la cosecha, y es fijo.
  - Las curvas por año se cortan donde cambiaría la composición del grupo.
  - La ventana de observación termina en jun-2026: desde jul-2026 hay un castigo
    que rompe la serie (ver LEEME.md).
"""
import json, re, sys, hashlib
from pathlib import Path
import numpy as np
import pandas as pd

AQUI = Path(__file__).resolve().parent
XLSX = Path(sys.argv[1]) if len(sys.argv) > 1 else AQUI / 'Cosechas_2026.xlsx'
SALIDA = AQUI / 'datos'
HTML = AQUI / 'Cosechas_Districolmotos.html'
CORTE_SERIE = pd.Period('2026-06', 'M')   # último mes comparable, antes del castigo
SALT = 'districolmotos-cosechas-2026'

MES = {'ENERO':1,'FEBRERO':2,'MARZO':3,'ABRIL':4,'MAYO':5,'JUNIO':6,'JULIO':7,
       'AGOSTO':8,'SEPTIEMBRE':9,'OCTUBRE':10,'NOVIEMBRE':11,'DICIEMBRE':12}
HOJAS_CORTE = {'FEBRERO 26':'2026-02','MARZO 26':'2026-03','ABRIL 26':'2026-04',
               'MAYO 26':'2026-05','JUNIO 26':'2026-06','JULIO 26':'2026-07','AGOSTO 26':'2026-08'}


# ---------------------------------------------------------------- cortes mensuales
def _col(df, *nombres):
    """Las hojas de febrero y abril traen otro layout: hay que buscar por alias."""
    cols = {str(c).strip().upper(): c for c in df.columns}
    for n in nombres:
        if n.strip().upper() in cols:
            return df[cols[n.strip().upper()]]
    return pd.Series([np.nan] * len(df), index=df.index)


def cortes_mensuales():
    out = []
    for hoja, corte in HOJAS_CORTE.items():
        df = pd.read_excel(XLSX, sheet_name=hoja)
        t = pd.DataFrame({
            'corte': corte,
            'pagare': _col(df, 'PAGARE').astype(str).str.strip(),
            'cedula': _col(df, 'CEDULASOCI').astype(str).str.strip(),
            'agencia': _col(df, 'AGENCIA').astype(str).str.strip(),
            'f_desemb': pd.to_datetime(_col(df, 'F_INICIOFI', 'FECHADESEM'), errors='coerce'),
            'saldo_cap': pd.to_numeric(_col(df, 'SALDOCAPIT'), errors='coerce'),
            # ojo: en la hoja de abril la columna trae un espacio al final
            'saldo_mora': pd.to_numeric(_col(df, 'SALDO EN MORA', 'SALDO EN MORA ', 'SALDO MORA'), errors='coerce'),
            'dias_mora': pd.to_numeric(_col(df, 'DIASMORA'), errors='coerce'),
            'cuotas_mora': pd.to_numeric(_col(df, 'CUOTASMORA'), errors='coerce'),
            'plazo': pd.to_numeric(_col(df, 'PLAZO'), errors='coerce'),
            'tasa': pd.to_numeric(_col(df, 'TASACOLOCA'), errors='coerce'),
            'valor_gar': pd.to_numeric(_col(df, 'VALORGARAN'), errors='coerce'),
            'cap_ini': pd.to_numeric(_col(df, 'CAPITALINI'), errors='coerce'),
            'categoria': _col(df, 'CATEGORIAA').astype(str).str.strip(),
            'linea': _col(df, 'NOMBRELINE').astype(str).str.strip(),
            'empresa': _col(df, 'NOMBREEMPR').astype(str).str.strip(),
            'garantia': _col(df, 'NOMBREGARA').astype(str).str.strip(),
            'asesor': _col(df, 'NOMBREASES').astype(str).str.strip(),
            'destino': _col(df, 'NOMBREDEST').astype(str).str.strip(),
            'fecha_reest': pd.to_datetime(_col(df, 'FECHAREEST'), errors='coerce'),
            'nro_reest': pd.to_numeric(_col(df, 'NROREESTRU'), errors='coerce'),
        })
        t['cosecha'] = t.f_desemb.dt.to_period('M').astype(str)
        out.append(t)
    return pd.concat(out, ignore_index=True)


# ---------------------------------------------------- triángulo (Consolidado + Cosechas)
def triangulo():
    """Consolidado manda (precisión completa); la hoja Cosechas solo rellena huecos."""
    import openpyxl
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)

    r = list(wb['Consolidado'].iter_rows(values_only=True))
    bloques = []
    for j, v in enumerate(r[2]):
        if v and str(v).strip().upper() == 'COSECHA':
            etiqueta = next((str(r[1][k]).strip() for k in range(j, -1, -1)
                             if k < len(r[1]) and r[1][k]), None)
            bloques.append((j, etiqueta))
    con = pd.DataFrame([
        {'cosecha': str(int(fila[j])), 'capital_mora': fila[j+1],
         'colocacion': fila[j+2], 'pct': fila[j+3],
         'obs': f"{lbl.split()[1]}-{MES[lbl.split()[0]]:02d}"}
        for j, lbl in bloques for fila in r[3:]
        if j < len(fila) and isinstance(fila[j], (int, float))
    ])

    rows = list(wb['Cosechas'].iter_rows(values_only=True))
    anio, mes, obs, cur = rows[0], rows[1], [], None
    for j in range(len(mes)):
        if j < len(anio) and anio[j] is not None:
            cur = int(anio[j])
        m = str(mes[j]).strip().upper() if mes[j] else ''
        obs.append(f'{cur}-{MES[m]:02d}' if m in MES and cur else None)
    tri = pd.DataFrame([
        {'cosecha': str(int(fila[0])), 'obs': o, 'pct': float(fila[j])}
        for fila in rows[2:] if fila[0] is not None
        for j, o in enumerate(obs)
        if o and j < len(fila) and isinstance(fila[j], (int, float))
    ])

    con['k'] = con.cosecha + '|' + con.obs
    tri['k'] = tri.cosecha + '|' + tri.obs
    base = pd.concat([con[['cosecha', 'obs', 'pct']],
                      tri.loc[~tri.k.isin(set(con.k)), ['cosecha', 'obs', 'pct']]], ignore_index=True)
    colo = con.groupby('cosecha').colocacion.max()
    base['colocacion'] = base.cosecha.map(colo)
    base['capital'] = base.pct * base.colocacion
    base['mob'] = (pd.PeriodIndex(base.obs, freq='M')
                   - pd.PeriodIndex(base.cosecha.str[:4] + '-' + base.cosecha.str[4:], freq='M')
                   ).map(lambda x: x.n)
    base['op'] = pd.PeriodIndex(base.obs, freq='M')
    base['anio'] = base.cosecha.str[:4]
    return base[base.mob >= 0].copy(), colo, con


# ---------------------------------------------------------------------- validaciones
def validar(cred, base, con, colo):
    print('\n— Controles —')
    v = colo.groupby(level=0).nunique() if colo.index.duplicated().any() else None
    n_var = int((con.groupby('cosecha').colocacion.nunique() > 1).sum())
    print(f'  [{"OK " if n_var == 0 else "!! "}] colocación constante entre bloques de observación '
          f'({n_var} cosechas variables)')

    ago = cred[(cred.corte == '2026-08') & (cred.cosecha >= '2022-01')]
    calc = ago.loc[ago.dias_mora > 30, 'saldo_mora'].sum()
    ref = con.loc[con.obs == '2026-08', 'capital_mora'].sum()
    print(f'  [{"OK " if abs(calc - ref) < 1 else "!! "}] numerador ago-2026: capital vencido con DIASMORA>30 '
          f'= {calc:,.0f} vs Consolidado {ref:,.0f}')
    print(f'  [i  ] exposición: el saldo TOTAL de esos créditos es '
          f'{ago.loc[ago.dias_mora > 30, "saldo_cap"].sum():,.0f} '
          f'({ago.loc[ago.dias_mora > 30, "saldo_cap"].sum() / calc:.1f}x el capital vencido)')

    jun = cred[cred.corte == '2026-06']; jul = cred[cred.corte == '2026-07']
    salio = jun.saldo_mora.sum() - jul.saldo_mora.sum()
    print(f'  [i  ] castigo jun→jul 2026: salieron COP {salio:,.0f} de mora '
          f'({len(set(jun.pagare) - set(jul.pagare))} créditos); '
          f'desembolso más antiguo pasó de {jun.f_desemb.min():%Y-%m} a {jul.f_desemb.min():%Y-%m}')


# ------------------------------------------------------------------------ agregados
def construir(cred, base, colo):
    pre = base[base.op <= CORTE_SERIE].dropna(subset=['pct', 'capital', 'colocacion'])

    # curvas por año, ponderadas y con COMPOSICIÓN CONSTANTE
    curvas, comp = {}, {}
    for y, g in pre.groupby('anio'):
        N = g.cosecha.nunique()
        ser = g.groupby('mob').apply(lambda d: d.capital.sum() / d.colocacion.sum() * 100,
                                     include_groups=False)
        nc = g.groupby('mob').cosecha.nunique()
        pts = [{'mob': int(m), 'pct': round(float(v), 3), 'n': int(nc[m])}
               for m, v in ser.items() if nc[m] == N and m <= 36 and not pd.isna(v)]
        if len(pts) >= 2:
            curvas[y] = pts
            comp[y] = {'N': int(N), 'mob_min': pts[0]['mob'], 'mob_max': pts[-1]['mob']}

    piv = (pre.pivot_table(index='cosecha', columns='mob', values='pct') * 100).reindex(sorted(colo.index))
    gv = lambda c, m: None if m not in piv.columns or pd.isna(piv.at[c, m]) else round(float(piv.at[c, m]), 2)
    mob_fijo = [{'cosecha': c, 'colocacion': float(colo[c]),
                 'm6': gv(c, 6), 'm12': gv(c, 12), 'm18': gv(c, 18)} for c in sorted(colo.index)]

    heat = [{'cosecha': r.cosecha, 'mob': int(r.mob), 'pct': round(r.pct * 100, 2)}
            for r in pre[pre.mob <= 36].itertuples() if not pd.isna(r.pct)]

    j6 = base[base.obs == '2026-06'].set_index('cosecha')
    j7 = base[base.obs == '2026-07'].set_index('cosecha')
    castigo = []
    for c in sorted(set(j6.index) & set(j7.index)):
        p6, p7, k6, k7 = j6.at[c, 'pct'], j7.at[c, 'pct'], j6.at[c, 'capital'], j7.at[c, 'capital']
        if any(pd.isna(x) for x in (p6, p7, k6, k7)):
            continue
        castigo.append({'cosecha': c, 'jun': round(p6 * 100, 2), 'jul': round(p7 * 100, 2),
                        'delta': round((p7 - p6) * 100, 2), 'monto': round(float(k6 - k7))})

    # segmentos: cosechas recientes, denominador = capital inicial (ver salvedad en LEEME)
    a = cred[cred.corte == '2026-08'].copy()
    a['punto'] = a.linea.str.replace('CARTERA CREDITOS', '', regex=False).str.split('CATEGORIA').str[0].str.strip()
    rec = a[a.cosecha >= '2025-01'].copy()
    rec['den'] = rec.cap_ini
    rec['num'] = np.where(rec.dias_mora > 30, rec.cap_ini, 0)
    rec['pzo'] = pd.cut(rec.plazo, [0, 12, 18, 24, 36, 99],
                        labels=['Hasta 12m', '13-18m', '19-24m', '25-36m', 'Mas de 36m'])

    def seg(col, lbl, ordenado=False):
        t = rec.groupby(col, observed=True).agg(n=('pagare', 'size'), den=('den', 'sum'), num=('num', 'sum'))
        t = t[t.n >= 30]                                   # n<30: no se reporta tasa
        if not ordenado:
            t = t.assign(_p=t.num / t.den).sort_values('_p', ascending=False)
        return {'label': lbl, 'ordenado': ordenado,
                'items': [{'k': str(i), 'n': int(r.n), 'pct': round(r.num / r.den * 100, 2),
                           'den': float(r.den)} for i, r in t.iterrows()]}

    segmentos = [seg('punto', 'Punto de venta'), seg('garantia', 'Tipo de garantía'),
                 seg('pzo', 'Plazo', ordenado=True), seg('empresa', 'Tipo de deudor')]

    kpi = {'colocacion_total': float(colo.sum()), 'n_cosechas': int(len(colo)),
           'castigo_monto': float(sum(x['monto'] for x in castigo if x['delta'] < -0.5)),
           'castigo_cosechas': int(sum(1 for x in castigo if x['delta'] < -1)),
           'saldo_ago': float(a.saldo_cap.sum()),
           'mora30_ago': float(a[a.dias_mora > 30].saldo_cap.sum()),
           'n_creditos_ago': int(len(a))}
    p12 = piv[12] if 12 in piv.columns else None
    for y in ['2022', '2023', '2024', '2025', '2026']:
        p18 = [x for x in curvas.get(y, []) if x['mob'] == 18]
        kpi[f'm18_{y}'] = p18[0]['pct'] if p18 else None
        sub = [c for c in colo.index if c.startswith(y) and p12 is not None and not pd.isna(p12.get(c, np.nan))]
        if sub:
            kpi[f'm12_{y}'] = round(sum(p12[c] / 100 * colo[c] for c in sub) / sum(colo[c] for c in sub) * 100, 3)
            kpi[f'm12_{y}_n'] = len(sub)
        else:
            kpi[f'm12_{y}'], kpi[f'm12_{y}_n'] = None, 0

    return {'curvas': curvas, 'comp': comp, 'mob_fijo': mob_fijo, 'heatmap': heat,
            'castigo': castigo, 'segmentos': segmentos,
            'colocacion': [{'cosecha': c, 'monto': float(colo[c])} for c in sorted(colo.index)],
            'kpi': kpi}


# -------------------------------------------------------------------------- salidas
def exportar(cred, d):
    SALIDA.mkdir(exist_ok=True)
    pd.DataFrame(d['mob_fijo']).rename(columns={
        'm6': 'mora30_mob6_pct', 'm12': 'mora30_mob12_pct',
        'm18': 'mora30_mob18_pct', 'colocacion': 'colocacion_cop'}
    ).to_csv(SALIDA / 'cosechas_por_mob.csv', index=False)
    pd.DataFrame(d['heatmap']).rename(columns={'pct': 'mora30_pct'}
        ).to_csv(SALIDA / 'triangulo_mob.csv', index=False)
    pd.DataFrame(d['castigo']).rename(columns={
        'jun': 'mora30_jun2026_pct', 'jul': 'mora30_jul2026_pct',
        'delta': 'delta_pp', 'monto': 'mora_que_salio_cop'}
    ).to_csv(SALIDA / 'castigo_julio_2026.csv', index=False)
    pd.DataFrame([{'dimension': s['label'], 'segmento': i['k'], 'creditos': i['n'],
                   'capital_originado_cop': i['den'], 'mora30_pct': i['pct']}
                  for s in d['segmentos'] for i in s['items']]
                 ).to_csv(SALIDA / 'segmentos.csv', index=False)
    pd.DataFrame([{'anio_cosecha': y, 'mob': p['mob'], 'mora30_pct': p['pct'],
                   'cosechas_observadas': p['n']}
                  for y, v in d['curvas'].items() for p in v]
                 ).to_csv(SALIDA / 'curvas_por_anio.csv', index=False)

    # base crédito a crédito, con el documento anonimizado
    x = cred.copy()
    x['deudor_id'] = x.cedula.map(lambda c: hashlib.sha256((SALT + str(c)).encode()).hexdigest()[:12])
    x = x.drop(columns=['cedula'])
    cols = ['corte', 'pagare', 'deudor_id'] + [c for c in x.columns if c not in ('corte', 'pagare', 'deudor_id')]
    x[cols].to_csv(SALIDA / 'cortes_mensuales_homologados.csv', index=False)

    if HTML.exists():
        h = HTML.read_text()
        nuevo = 'const D = ' + json.dumps(d, ensure_ascii=False) + ';'
        h2 = re.sub(r'const D = \{.*?\};', lambda _: nuevo, h, count=1, flags=re.S)
        if h2 != h:
            HTML.write_text(h2)
            print(f'\n  tablero actualizado: {HTML.name}')


if __name__ == '__main__':
    if not XLSX.exists():
        sys.exit(f'No encuentro el Excel: {XLSX}\nUso: python3 analizar_cosechas.py ruta/al/archivo.xlsx')
    print(f'Leyendo {XLSX.name} …')
    cred = cortes_mensuales()
    base, colo, con = triangulo()
    print(f'  {len(cred):,} filas crédito-corte · {len(colo)} cosechas · '
          f'COP {colo.sum():,.0f} colocados')
    validar(cred, base, con, colo)
    d = construir(cred, base, colo)
    exportar(cred, d)
    print(f'\n  tablas en {SALIDA}/')
    k = d['kpi']
    print(f"\n— Mora 30+ al mes 12, por año de cosecha —")
    for y in ['2022', '2023', '2024', '2025']:
        print(f"    {y}: {k[f'm12_{y}']:.2f}%   ({k[f'm12_{y}_n']} cosechas)")
