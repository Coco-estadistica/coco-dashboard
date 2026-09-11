#!/usr/bin/env python3
"""
Prepara los insumos del informe a junta: deck_data.json y heatmap.png.

Lee las tablas que deja ../analizar_cosechas.py en ../datos/ y produce todo lo que
build_deck.js necesita. Correr en este orden:

    python3 ../analizar_cosechas.py Cosechas_2026.xlsx
    python3 preparar_datos.py
    node build_deck.js

Reglas que respeta:
  - Un dato faltante es s/d, nunca cero.
  - Las cosechas solo se comparan a MOB constante.
  - Los cortes por segmento solo usan los meses que traen CAPITALINI (marzo, mayo,
    junio, julio y agosto de 2026; febrero y abril vienen con otro layout y sin esa
    columna), y se miden al MOB 6, donde la cobertura sobre lo colocado es del 81%.
"""
import json
from pathlib import Path

import numpy as np
import pandas as pd

AQUI = Path(__file__).resolve().parent
DATOS = AQUI.parent / 'datos'
CORTES_CON_CAPITAL_INICIAL = ['2026-03', '2026-05', '2026-06', '2026-07', '2026-08']
BALDES = ['Al dia', '1-30', '31-60', '61-90', '91-180', '+180']
CORTES_ROLL = ['2026-02', '2026-03', '2026-04', '2026-05', '2026-06']   # pre-castigo
MOB_SEGMENTO = 6


def cargar():
    tri = pd.read_csv(DATOS / 'triangulo_mob.csv', dtype={'cosecha': str})
    colo = pd.read_csv(DATOS / 'cosechas_por_mob.csv', dtype={'cosecha': str}).set_index('cosecha')
    cur = pd.read_csv(DATOS / 'curvas_por_anio.csv', dtype={'anio_cosecha': str})
    cred = pd.read_csv(DATOS / 'cortes_mensuales_homologados.csv')
    cas = pd.read_csv(DATOS / 'castigo_julio_2026.csv', dtype={'cosecha': str})
    return tri, colo, cur, cred, cas


def roll_rates(cred):
    """Matriz de transición mes a mes, ponderada por saldo. Excluye jun→jul (castigo)."""
    d = cred.copy()
    d['b'] = pd.cut(d.dias_mora, [-1, 0, 30, 60, 90, 180, 99999], labels=BALDES)
    acc = None
    for t0, t1 in zip(CORTES_ROLL[:-1], CORTES_ROLL[1:]):
        x = d[d.corte == t0][['pagare', 'b', 'saldo_cap']]
        y = d[d.corte == t1][['pagare', 'b']]
        m = x.merge(y, on='pagare', suffixes=('_0', '_1'), how='left')
        m['b_1'] = m.b_1.astype(object).fillna('Sale')      # se pagó o salió del reporte
        tt = m.pivot_table(index='b_0', columns='b_1', values='saldo_cap',
                           aggfunc='sum', observed=True).fillna(0)
        acc = tt if acc is None else acc.add(tt, fill_value=0)
    r = (acc.div(acc.sum(axis=1), axis=0) * 100).reindex(BALDES).round(1)
    filas = [{'desde': b, 'cura': float(r.loc[b, 'Al dia']),
              'igual': float(r.loc[b, b]) if b in r.columns else 0.0,
              'empeora': float(sum(r.loc[b, c] for c in BALDES[BALDES.index(b) + 1:]
                                   if c in r.columns))} for b in BALDES[:-1]]
    return filas, float(r.loc['Al dia', '1-30'])


def por_punto_a_edad_constante(cred):
    """Mora 30+ por punto de venta midiendo todas las cosechas a la misma edad."""
    v = cred[cred.corte.isin(CORTES_CON_CAPITAL_INICIAL)].copy()
    v['punto'] = (v.linea.str.replace('CARTERA CREDITOS', '', regex=False)
                  .str.split('CATEGORIA').str[0].str.strip())
    v['mob'] = (pd.PeriodIndex(v.corte, freq='M')
                - pd.PeriodIndex(v.cosecha, freq='M')).map(lambda z: z.n)
    v['num'] = np.where(v.dias_mora > 30, v.cap_ini, 0)
    s = v[v.mob == MOB_SEGMENTO]
    g = s.groupby('punto').agg(n=('pagare', 'size'), den=('cap_ini', 'sum'), num=('num', 'sum'))
    g = g[g.n >= 20]                                        # n<20: no se reporta tasa
    limpio = lambda p: p.replace('SUZUKI- ', 'SUZUKI ').replace(' -', ' ').title()
    items = [{'punto': limpio(i), 'n': int(x.n), 'pct': round(x.num / x.den * 100, 1),
              'orig_M': round(x.den / 1e6)}
             for i, x in g.sort_values('num', ascending=False).iterrows()]
    return items, round(float(s.num.sum() / s.cap_ini.sum() * 100), 1), sorted(s.cosecha.unique())


def construir():
    tri, colo, cur, cred, cas = cargar()
    out = {}

    out['cosechas'] = {c: {'mob': [int(x) for x in g.sort_values('mob').mob],
                           'pct': [round(float(x), 2) for x in g.sort_values('mob').mora30_pct],
                           'colocacion': float(colo.colocacion_cop[c]),
                           'm12': None if pd.isna(colo.mora30_mob12_pct[c])
                                  else float(colo.mora30_mob12_pct[c])}
                       for c, g in tri.groupby('cosecha')}
    out['curvas_anio'] = {y: {'mob': [int(x) for x in g.mob],
                              'pct': [float(x) for x in g.mora30_pct]}
                          for y, g in cur.groupby('anio_cosecha')}
    out['mob_fijo'] = [{'cosecha': i, 'colocacion': float(r.colocacion_cop),
                        'm6': None if pd.isna(r.mora30_mob6_pct) else float(r.mora30_mob6_pct),
                        'm12': None if pd.isna(r.mora30_mob12_pct) else float(r.mora30_mob12_pct)}
                       for i, r in colo.iterrows()]
    out['colocacion'] = [{'cosecha': i, 'monto_M': round(r.colocacion_cop / 1e6)}
                         for i, r in colo.iterrows()]

    out['roll'], out['entrada_mora'] = roll_rates(cred)
    out['punto_mob6'], out['punto_mob6_total'], out['punto_mob6_cosechas'] = por_punto_a_edad_constante(cred)

    a = cred[cred.corte == '2026-08'].copy()
    a['mora_exp'] = np.where(a.dias_mora > 30, a.saldo_cap, 0)     # exposición, no vencido
    b = pd.cut(a.dias_mora, [-1, 0, 30, 60, 90, 180, 360, 99999],
               labels=['Al día', '1-30', '31-60', '61-90', '91-180', '181-360', '+360 días'])
    t = a.groupby(b, observed=True).agg(n=('pagare', 'size'), saldo=('saldo_cap', 'sum'))
    out['altura'] = [{'balde': str(i), 'n': int(x.n), 'saldo_M': round(x.saldo / 1e6),
                      'pct': round(x.saldo / t.saldo.sum() * 100, 1)} for i, x in t.iterrows()]
    out['cartera'] = {'saldo_M': round(a.saldo_cap.sum() / 1e6), 'n': int(len(a)),
                      'vencido_M': round(a[a.dias_mora > 30].saldo_mora.sum() / 1e6),
                      'exposicion_M': round(a.mora_exp.sum() / 1e6),
                      'n_mora': int((a.dias_mora > 30).sum())}

    x = a.sort_values('mora_exp', ascending=False)
    x['ac'] = x.mora_exp.cumsum() / x.mora_exp.sum() * 100
    k = int(np.ceil(len(x) * 0.10))
    out['concentracion'] = {'creditos': k, 'pct_mora': round(float(x.ac.iloc[k - 1]), 1),
                            'monto_M': round(float(x.mora_exp.head(k).sum() / 1e6)),
                            'total_M': round(float(a.mora_exp.sum() / 1e6))}

    a['anio'] = a.cosecha.str[:4]
    e = a.groupby('anio').agg(saldo=('saldo_cap', 'sum'), mora=('mora_exp', 'sum'))
    out['exposicion'] = [{'anio': i, 'saldo_M': round(x.saldo / 1e6),
                          'pct': round(x.saldo / e.saldo.sum() * 100, 1),
                          'icv': round(x.mora / x.saldo * 100, 1)}
                         for i, x in e.iterrows() if x.saldo > 1e7]

    c = cas[cas.delta_pp < -1].sort_values('delta_pp')
    out['castigo'] = [{'cosecha': r.cosecha, 'jun': float(r.mora30_jun2026_pct),
                       'jul': float(r.mora30_jul2026_pct),
                       'monto_M': round(r.mora_que_salio_cop / 1e6)} for r in c.itertuples()]
    return out


def heatmap(d):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.colors import LinearSegmentedColormap

    cos, MOB = sorted(d['cosechas']), 37
    M = np.full((len(cos), MOB), np.nan)                  # el vacío es s/d, no cero
    for i, c in enumerate(cos):
        for m, p in zip(d['cosechas'][c]['mob'], d['cosechas'][c]['pct']):
            if m < MOB:
                M[i, m] = p
    cmap = LinearSegmentedColormap.from_list('c', ['#E8F1FD', '#9EC5F4', '#2A78D6', '#1C5CAB', '#0D366B'])
    cmap.set_bad('#FFFFFF')
    fig, ax = plt.subplots(figsize=(12.4, 6.0), dpi=200)
    im = ax.imshow(M, cmap=cmap, vmin=0, vmax=12, aspect='auto', interpolation='nearest')
    ax.set_xticks(range(0, MOB, 6)); ax.set_xticklabels(range(0, MOB, 6), fontsize=9, color='#62717F')
    mesn = {'01': 'ene', '07': 'jul'}
    yt = [i for i, c in enumerate(cos) if c[4:] in mesn]
    ax.set_yticks(yt)
    ax.set_yticklabels([f'{mesn[cos[i][4:]]}-{cos[i][:4]}' for i in yt], fontsize=9, color='#62717F')
    ax.set_xlabel('Meses en libros desde el desembolso (MOB)', fontsize=10, color='#62717F', labelpad=8)
    for s in ax.spines.values():
        s.set_visible(False)
    ax.tick_params(length=0)
    ax.set_xticks(np.arange(-.5, MOB, 1), minor=True)
    ax.set_yticks(np.arange(-.5, len(cos), 1), minor=True)
    ax.grid(which='minor', color='white', linewidth=1.1)
    ax.tick_params(which='minor', length=0)
    cb = fig.colorbar(im, ax=ax, fraction=.022, pad=.015)
    cb.set_label('% de la colocación en mora 30+', fontsize=9, color='#62717F')
    cb.ax.tick_params(labelsize=8, length=0, colors='#62717F')
    cb.set_ticks([0, 3, 6, 9, 12])
    cb.ax.set_yticklabels(['0%', '3%', '6%', '9%', '12% o más'])
    cb.outline.set_visible(False)
    plt.tight_layout()
    plt.savefig(AQUI / 'heatmap.png', bbox_inches='tight', facecolor='white')


if __name__ == '__main__':
    d = construir()
    json.dump(d, open(AQUI / 'deck_data.json', 'w'), ensure_ascii=False)
    heatmap(d)
    k = d['cartera']
    print(f"deck_data.json — {len(d['cosechas'])} cosechas, {len(d['punto_mob6'])} puntos de venta")
    print(f"heatmap.png — listo")
    print(f"cartera ago-2026: {k['n']} créditos, COP {k['saldo_M']:,} M · "
          f"vencido COP {k['vencido_M']} M · exposición COP {k['exposicion_M']} M")
