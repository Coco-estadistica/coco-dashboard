#!/usr/bin/env python3
"""
Analítica de riesgo para el comité de crédito y cartera — Districolmotos.

Lee las tablas que deja ../analizar_cosechas.py en ../datos/ y produce datos_comite.json,
que alimenta build_comite.js.

    python3 ../analizar_cosechas.py Cosechas_2026.xlsx
    python3 preparar_comite.py
    node build_comite.js

Requiere pandas, numpy, scipy y statsmodels.

Decisiones metodológicas que NO se deben cambiar sin pensarlo:

  - Se excluye siempre la transición jun→jul 2026: el castigo de julio la distorsiona.
  - Se excluyen los créditos vivos pasados de su plazo pactado. Siguen en el libro
    precisamente porque no pagaron (93% en mora): incluirlos fabrica una relación
    entre plazo y riesgo que no existe.
  - La regresión usa un solo corte por crédito, para no contar siete veces el mismo caso.
  - Ninguna tasa se reporta con menos de 30 observaciones (20 en las curvas de cura).
  - Un dato faltante es s/d, nunca cero.
"""
import json
from pathlib import Path

import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
from scipy import stats

AQUI = Path(__file__).resolve().parent
DATOS = AQUI.parent / 'datos'
# pares (t0, t1) de cortes consecutivos; falta jun→jul a propósito
PARES = [('2026-02', '2026-03'), ('2026-03', '2026-04'), ('2026-04', '2026-05'),
         ('2026-05', '2026-06'), ('2026-07', '2026-08')]
BALDES = ['Al dia', '1-30', '31-60', '61-90', '91-180', '+180']
CAL = ['A', 'B', 'C', 'D', 'E']
CORTE_ACTUAL = '2026-08'
pc = lambda v, d=1: f'{v:.{d}f}'.replace('.', ',')      # coma decimal, como el resto del informe


def cargar():
    d = pd.read_csv(DATOS / 'cortes_mensuales_homologados.csv')
    d['mob'] = (pd.PeriodIndex(d.corte, freq='M')
                - pd.PeriodIndex(d.cosecha, freq='M')).map(lambda x: x.n)
    d['m30'] = (d.dias_mora > 30).astype(int)
    d['punto'] = (d.linea.str.replace('CARTERA CREDITOS', '', regex=False)
                  .str.split('CATEGORIA').str[0].str.strip())
    return d


def transiciones(d, filtro):
    """Empareja cada crédito entre cortes consecutivos. `inner`: los que salen del libro
    no entran al denominador, se miden aparte en prepago()."""
    out = []
    for t0, t1 in PARES:
        s0 = d[d.corte == t0]
        x = s0[filtro(s0)][['pagare', 'mob', 'saldo_cap', 'dias_mora', 'plazo']]
        y = d[d.corte == t1][['pagare', 'm30', 'dias_mora']]
        out.append(x.merge(y, on='pagare', how='inner', suffixes=('_0', '_1')))
    return pd.concat(out)


def hazard(d):
    """De los que estaban al día (0 días), cuántos entran en mora al mes siguiente."""
    h = transiciones(d, lambda s: s.dias_mora == 0)
    h['b'] = pd.cut(h.mob, [-1, 2, 5, 8, 11, 14, 17, 23, 99],
                    labels=['0-2', '3-5', '6-8', '9-11', '12-14', '15-17', '18-23', '24+'])
    g = h.groupby('b', observed=True).agg(n=('pagare', 'size'),
                                          entra=('dias_mora_1', lambda s: (s > 0).sum()))
    return [{'edad': str(i), 'n': int(r.n), 'pct': round(r.entra / r.n * 100, 1)}
            for i, r in g.iterrows() if r.n >= 30]


def cura(d):
    """De los que estaban en mora 30+, cuántos bajan de 30 días al mes siguiente."""
    c = transiciones(d, lambda s: s.dias_mora > 30)
    c['b'] = pd.cut(c.dias_mora_0, [30, 60, 90, 180, 360, 99999],
                    labels=['31-60', '61-90', '91-180', '181-360', '+360'])
    g = c.groupby('b', observed=True).agg(n=('pagare', 'size'),
                                          sale=('dias_mora_1', lambda s: (s <= 30).sum()))
    return [{'altura': str(i), 'n': int(r.n), 'pct': round(r.sale / r.n * 100, 1)}
            for i, r in g.iterrows() if r.n >= 20]


def prepago(d):
    """Créditos sanos que desaparecen del corte siguiente: pagaron o cancelaron."""
    out = []
    for t0, t1 in PARES:
        x = d[d.corte == t0][['pagare', 'mob', 'dias_mora']]
        vivos = set(d[d.corte == t1].pagare)
        out.append(x.assign(sale=~x.pagare.isin(vivos)))
    s = pd.concat(out)
    s['b'] = pd.cut(s.mob, [-1, 5, 11, 17, 23, 99],
                    labels=['0-5', '6-11', '12-17', '18-23', '24+'])
    g = s[s.dias_mora <= 30].groupby('b', observed=True).agg(n=('pagare', 'size'),
                                                            sale=('sale', 'mean'))
    return [{'edad': str(i), 'n': int(r.n), 'pct': round(r.sale * 100, 1)}
            for i, r in g.iterrows() if r.n >= 30]


def migracion(d):
    """A dónde va el saldo de cada calificación al mes siguiente."""
    acc = None
    for t0, t1 in PARES:
        x = d[d.corte == t0][['pagare', 'categoria', 'saldo_cap']]
        y = d[d.corte == t1][['pagare', 'categoria']]
        m = x.merge(y, on='pagare', suffixes=('_0', '_1'), how='left')
        m['categoria_1'] = m.categoria_1.fillna('Sale')
        tt = m.pivot_table(index='categoria_0', columns='categoria_1',
                           values='saldo_cap', aggfunc='sum').fillna(0)
        acc = tt if acc is None else acc.add(tt, fill_value=0)
    mg = (acc.div(acc.sum(axis=1), axis=0) * 100).reindex(CAL)
    return [{'desde': k, **{c: round(float(mg.loc[k, c]), 1) for c in CAL if c in mg.columns},
             'Sale': round(float(mg.loc[k, 'Sale']), 1)} for k in CAL]


def discrimina(d):
    """¿Alguna característica del crédito separa buenos de malos, controlando la edad?"""
    a = d[d.corte == CORTE_ACTUAL].copy()
    lim = a[(a.cosecha >= '2024-01') & (a.mob >= 3) & (a.mob < a.plazo)].copy()
    lim = lim[lim.groupby('punto').punto.transform('size') >= 30]
    lim['plazo_g'] = pd.cut(lim.plazo, [0, 12, 18, 24, 99],
                            labels=['Hasta 12m', '13-18m', '19-24m', 'Más de 24m']).astype(str)
    M = {k: smf.logit(v, data=lim).fit(disp=0) for k, v in {
        'completo': 'm30 ~ mob + C(punto) + C(plazo_g)',
        'sin_punto': 'm30 ~ mob + C(plazo_g)',
        'sin_plazo': 'm30 ~ mob + C(punto)'}.items()}
    lrt = lambda red: stats.chi2.sf(2 * (M['completo'].llf - M[red].llf),
                                    int(M['completo'].df_model - M[red].df_model))
    filas = [
        {'variable': 'Edad del crédito', 'p': round(float(M['completo'].pvalues['mob']), 4),
         'nota': f"OR {pc(np.exp(M['completo'].params['mob']), 2)} por mes"},
        {'variable': 'Punto de venta', 'p': round(float(lrt('sin_punto')), 4),
         'nota': '6 grados de libertad'},
        {'variable': 'Plazo', 'p': round(float(lrt('sin_plazo')), 4),
         'nota': '3 grados de libertad'},
    ]
    # las de abajo se prueban a edad constante (MOB 6), no en el modelo
    v = d[d.corte.isin(['2026-03', '2026-05', '2026-06', '2026-07', '2026-08'])]
    s6 = v[v.mob == 6].copy()

    def dos_grupos(mascara1, mascara2):
        x, y = s6[mascara1].m30, s6[mascara2].m30
        tab = np.array([[x.sum(), len(x) - x.sum()], [y.sum(), len(y) - y.sum()]])
        return float(stats.chi2_contingency(tab)[1]), x.mean() * 100, y.mean() * 100

    p, v1, v2 = dos_grupos(s6.tasa.round(1) == 3.5, s6.tasa.round(1) == 3.8)
    filas.append({'variable': 'Tasa (3,5% vs 3,8%)', 'p': round(p, 4),
                  'nota': f'{pc(v1)}% vs {pc(v2)}%'})
    p, v1, v2 = dos_grupos(s6.garantia == 'PRENDA', s6.garantia == 'PRENDA SIN CODEUDOR')
    filas.append({'variable': 'Codeudor', 'p': round(p, 4), 'nota': f'{pc(v1)}% vs {pc(v2)}%'})
    rho, pv = stats.spearmanr(s6.cap_ini, s6.m30)
    filas.append({'variable': 'Monto del crédito', 'p': round(float(pv), 4),
                  'nota': f'Spearman r = {pc(rho, 3)}'})
    return filas, {'n': int(len(lim)), 'mora': round(float(lim.m30.mean() * 100), 1)}


def concentracion(d):
    a = d[d.corte == CORTE_ACTUAL].copy()
    dd = a.groupby('deudor_id').saldo_cap.sum().sort_values(ascending=False)
    por_deudor = [{'top': p, 'pct': round(float(dd.head(int(np.ceil(len(dd) * p / 100))).sum()
                                                / dd.sum() * 100), 1)} for p in [1, 5, 10, 20]]
    a['mora_exp'] = np.where(a.dias_mora > 30, a.saldo_cap, 0)
    x = a.sort_values('mora_exp', ascending=False)
    x['ac'] = x.mora_exp.cumsum() / x.mora_exp.sum() * 100
    de_mora = [{'pct_creditos': p,
                'pct_mora': round(float(x.ac.iloc[int(np.ceil(len(x) * p / 100)) - 1]), 1)}
               for p in [2, 5, 10, 15, 20]]
    return por_deudor, de_mora


def proyeccion():
    """Chain-ladder sobre el pico: cuánto multiplica la mora del mes k hasta su máximo."""
    tri = pd.read_csv(DATOS / 'triangulo_mob.csv', dtype={'cosecha': str})
    colo = pd.read_csv(DATOS / 'cosechas_por_mob.csv',
                       dtype={'cosecha': str}).set_index('cosecha')
    piv = tri.pivot(index='cosecha', columns='mob', values='mora30_pct')
    madura = piv[piv.columns[piv.columns <= 30]].dropna(thresh=25)   # 25+ meses observados
    pico = madura.max(axis=1)
    fs = {k: (pico / madura[k]).replace([np.inf], np.nan).dropna().pipe(lambda s: s[s >= 1])
          for k in [6, 9, 12, 15, 18, 24]}
    factores = [{'mes': k, 'mediana': round(float(f.median()), 2),
                 'p25': round(float(f.quantile(.25)), 2), 'p75': round(float(f.quantile(.75)), 2),
                 'n': int(len(f))} for k, f in fs.items()]
    proy, joven = [], []
    for c in piv.index:
        s = piv.loc[c].dropna()
        mm = int(s.index.max())
        if mm >= 25:
            continue
        ks = [x for x in fs if x <= mm and not pd.isna(piv.loc[c, x])]
        if not ks:      # demasiado joven: el factor del mes 3 va de 8,7x a 35x, no dice nada
            joven.append({'cosecha': c, 'mob': mm,
                          'col_M': round(colo.colocacion_cop[c] / 1e6)})
            continue
        k = max(ks)
        act, f = float(piv.loc[c, k]), fs[k]
        proy.append({'cosecha': c, 'mob': mm, 'hoy': round(float(s.iloc[-1]), 2),
                     'proy': round(act * f.median(), 2), 'p25': round(act * f.quantile(.25), 2),
                     'p75': round(act * f.quantile(.75), 2),
                     'col_M': round(colo.colocacion_cop[c] / 1e6)})
    P = pd.DataFrame(proy)
    e = lambda col: float((P[col] / 100 * P.col_M).sum())
    total = {'colocacion_M': int(P.col_M.sum()), 'hoy_M': round(e('hoy')),
             'proy_M': round(e('proy')), 'p25_M': round(e('p25')), 'p75_M': round(e('p75')),
             'falta_M': round(e('proy') - e('hoy')), 'n_cosechas': len(P),
             'jovenes_M': int(sum(x['col_M'] for x in joven)), 'n_jovenes': len(joven)}
    pico_info = {'mediana': round(float(pico.median()), 2), 'min': round(float(pico.min()), 2),
                 'max': round(float(pico.max()), 2), 'n': int(len(pico))}
    return factores, proy, joven, total, pico_info


def maduracion(d, haz):
    """Cruza el saldo al día por edad con la curva de hazard: cuánta mora nueva viene."""
    a = d[d.corte == CORTE_ACTUAL].copy()
    a['b'] = pd.cut(a.mob, [-1, 2, 5, 8, 11, 14, 17, 23, 999],
                    labels=['0-2', '3-5', '6-8', '9-11', '12-14', '15-17', '18-23', '24+'])
    HZ = {z['edad']: z['pct'] for z in haz}
    g = a[a.dias_mora == 0].groupby('b', observed=True).agg(n=('pagare', 'size'),
                                                            saldo=('saldo_cap', 'sum'))
    filas = [{'edad': str(i), 'n': int(r.n), 'saldo_M': round(r.saldo / 1e6),
              'hazard': HZ.get(str(i)),
              'entra_M': round(r.saldo / 1e6 * HZ.get(str(i), 0) / 100, 1)}
             for i, r in g.iterrows()]
    total = {'saldo_M': round(g.saldo.sum() / 1e6),
             'entra_M': round(sum(x['entra_M'] for x in filas), 1),
             'pct_joven': round(float(g.saldo[['0-2', '3-5']].sum() / g.saldo.sum() * 100))}
    return filas, total


if __name__ == '__main__':
    d = cargar()
    out = {}
    out['hazard'] = hazard(d)
    out['cura'] = cura(d)
    out['prepago'] = prepago(d)
    out['migracion'] = migracion(d)
    out['discrimina'], out['muestra_logit'] = discrimina(d)
    out['conc_deudor'], out['conc_mora'] = concentracion(d)
    (out['factores'], out['proyeccion'], out['jovenes'],
     out['proy_total'], out['pico_maduras']) = proyeccion()
    out['maduracion'], out['maduracion_total'] = maduracion(d, out['hazard'])
    json.dump(out, open(AQUI / 'datos_comite.json', 'w'), ensure_ascii=False, indent=1)

    print('datos_comite.json listo\n')
    print('— Qué discrimina el riesgo —')
    for v in out['discrimina']:
        marca = 'SÍ' if v['p'] < 0.05 else ('marginal' if v['p'] < 0.10 else 'no')
        print(f"  {v['variable']:24s} p={v['p']:<8} {marca}")
    t = out['proy_total']
    print(f"\n— Provisión —\n  falta por aflorar: COP {t['falta_M']} M "
          f"(rango {t['p25_M']}–{t['p75_M']} M) sobre {t['n_cosechas']} cosechas inmaduras")
    print(f"\n— Cobranza —\n  cura: " +
          ' · '.join(f"{c['altura']}d {pc(c['pct'])}%" for c in out['cura']))
