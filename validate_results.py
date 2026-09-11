from pathlib import Path
import json, numpy as np
import pandas as pd
from openpyxl import load_workbook
ROOT=Path(__file__).resolve().parent; OUT=ROOT/'results'; FIG=ROOT/'figures'
summary=json.loads((OUT/'summary.json').read_text())
q1=summary['q1']; assert abs(q1['terminal_soc']-6000)<1e-6
assert 1200-1e-6<=q1['soc_min']<=10800+1e-6 and 1200-1e-6<=q1['soc_max']<=10800+1e-6
assert q1['purchase_kwh']>0
assert q1['dp_grid_value'] >= q1['objective']
assert 0 < q1['dp_relative_gap_pct'] < 1.0
assert (OUT/'value_dp_q1.json').exists()
diag=pd.read_csv(OUT/'forecast_diagnostics.csv') if (OUT/'forecast_diagnostics.csv').exists() else None
assert diag is not None and set(diag['series']) >= {'load_median7','load_level_shape','pv_median3','pv_median7'}
assert float(diag.loc[diag.series=='load_level_shape','mape_pct'].mean()) < float(diag.loc[diag.series=='load_median7','mape_pct'].mean())
assert float(diag.loc[diag.series=='pv_median3','mape_pct'].mean()) < float(diag.loc[diag.series=='pv_median7','mape_pct'].mean())
for fn,sheets in {'result1.xlsx':['计划购电量','充放电量'],'result2.xlsx':['计划购电量','充放电量','紧急购电量'],'result3.xlsx':['计划购电量','调整购电量','充放电量','紧急购电量'],'result4-2.xlsx':['计划购电量','充放电量','紧急购电量'],'result4-3.xlsx':['计划购电量','调整购电量','充放电量','紧急购电量']}.items():
 p=OUT/fn; assert p.exists() and p.stat().st_size>0
 wb=load_workbook(p,read_only=True,data_only=True); assert wb.sheetnames==sheets,(fn,wb.sheetnames)
 for ws in wb.worksheets: assert ws.max_row>=2 and ws.max_column>=2
# required figure coverage: 3 categories x q1..q4 (logical pngs)
for q in range(1,5):
 for cat in ('raw','process','result'):
  found=list(FIG.glob(f'{cat}_q{q}_*.png'))
  assert found,(cat,q)
print('PASS result validation')
print('Q1 purchase_kwh',q1['purchase_kwh'],'Q1 soc',q1['soc_min'],q1['soc_max'])
print('aggregates',summary['aggregate'])
df=pd.read_csv(OUT/'daily_metrics.csv')
assert np.allclose(df['q3_cost'],df['q3_base_cost']+df['q3_reduction_fee']+df['q3_increase_fee']+df['q3_emergency_cost'],rtol=0,atol=1e-6)
assert np.allclose(df['q4_3_cost'],df['q4_3_base_cost']+df['q4_3_reduction_fee']+df['q4_3_increase_fee']+df['q4_3_emergency_cost'],rtol=0,atol=1e-6)
assert (df[['q3_base_cost','q3_reduction_fee','q3_increase_fee','q4_3_base_cost','q4_3_reduction_fee','q4_3_increase_fee']] >= -1e-9).all().all()
print('fig_png',len(list(FIG.glob('*.png'))),'fig_svg',len(list(FIG.glob('*.svg'))))
