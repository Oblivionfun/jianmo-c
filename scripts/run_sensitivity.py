"""Small deterministic sensitivity run used in the paper, with no future data."""
from pathlib import Path
import sys
import json
import numpy as np
import pandas as pd
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import solve_c
from state_value import forecast_load_decomposed

root = Path(__file__).resolve().parents[1]
a1, load, pv, _, _ = solve_c.read_inputs()
rows = []
for factor in (0.80, 0.90, 1.00):
    soc = solve_c.S_INIT
    cost = emergency = 0.0
    for d in solve_c.EVAL_DATES:
        lh = forecast_load_decomposed(load, d, solve_c.DATES, solve_c.DT)
        gh = factor * solve_c.forecast_pv(pv, d, 3)
        r = solve_c.solve_dispatch(lh, gh, a1.price.values, initial=soc, emergency=True)
        e = solve_c.actual_emergency(load.loc[d].values, pv.loc[d].values,
                                     r['q'], r['c'], r['d'], r['u'])
        cost += float(np.dot(a1.price.values, r['q']) + np.dot(5 * a1.price.values, e))
        emergency += float(e.sum())
        soc = float(r['s'][-1])
    rows.append({'pv_factor': factor, 'q2_cost': cost, 'q2_emergency_kwh': emergency})
pd.DataFrame(rows).to_csv(root / 'results/sensitivity_forecast.csv', index=False)
manifest_path = root / 'results/复现清单.json'
if manifest_path.exists():
    manifest = json.loads(manifest_path.read_text(encoding='utf-8'))
    outputs = {str(x) for x in manifest.get('outputs', [])}
    outputs.add('results/sensitivity_forecast.csv')
    manifest['outputs'] = sorted(outputs)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
print(pd.DataFrame(rows).to_string(index=False))
