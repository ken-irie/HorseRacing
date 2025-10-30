import pandas as pd
from pathlib import Path
from datetime import datetime

def get_output_dir() -> Path:
    try:
        base = Path(__file__).resolve().parent
    except NameError:
        base = Path.cwd()
    out = base / "output"
    out.mkdir(parents=True, exist_ok=True)
    return out

df = pd.DataFrame({"A": [1,2,3], "B": ["x","y","z"]})
out_path = get_output_dir() / f"data_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
with pd.ExcelWriter(out_path, engine="openpyxl") as w:
    df.to_excel(w, index=False, sheet_name="Sheet1")

print(f"Excel出力: {out_path}")
