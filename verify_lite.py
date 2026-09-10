# -*- coding: utf-8 -*-
from collections import Counter
import formulas

xl = formulas.ExcelModel().loads("/Users/kk/workspaces/side-hustle/recon-pro/lite/ecom-recon-lite.xlsx").finish()
sol = xl.calculate()

def get(coord):
    suffix = "]自动对账'!" + coord
    for k, v in sol.items():
        if k.upper().endswith(suffix.upper()):
            try:
                return v.value[0, 0]
            except Exception:
                return v.value
    return None

c = Counter()
for r in range(2, 14):
    s = get(f"H{r}")
    if isinstance(s, str) and s:
        c[s] += 1
print(dict(c))
assert c["⚠未结算"] == 1 and c["❌金额差异"] == 1 and c["✓"] == 10, "MISMATCH"
print("LITE ENGINE PASS")
