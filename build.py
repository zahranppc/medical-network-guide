"""Build data.js from the Medright Medical Network Excel file."""
import openpyxl, json, re, sys, warnings
warnings.filterwarnings("ignore")
SRC = sys.argv[1] if len(sys.argv) > 1 else "/Users/zahran/Desktop/3,Medical Network Guide (3) (1) (1).xlsx"
TIERS = ["Elite", "Premium", "Classic", "Regular", "Basic"]
wb = openpyxl.load_workbook(SRC, read_only=True)
rows = {}
for ti, t in enumerate(TIERS):
    for i, row in enumerate(wb[t].iter_rows(values_only=True)):
        if i < 2: continue
        r = [re.sub(r"\s+", " ", str(v)).strip() if v is not None else "" for v in row[:7]]
        if not r[0]: continue
        rows.setdefault(tuple(r[:6]), [*r, 0])[7] |= 1 << ti
out = sorted(rows.values(), key=lambda r: (r[3], r[4], r[0].lower()))
json.dump(out, open("data.js", "w"), ensure_ascii=False, separators=(",", ":"))
with open("data.js", "r+") as f:
    body = f.read(); f.seek(0); f.write("window.DATA=" + body + ";")
print(len(out), "providers written")
assert len(out) > 5000 and all(r[7] for r in out)
