from datetime import datetime
import pandas as pd


def parse_japanese_datetime(s):
    if not isinstance(s, str):
        return pd.NaT
    s = s.replace("ごろ", "").strip()
    try:
        return datetime.strptime(s, "%Y年%m月%d日 %H時%M分")
    except Exception:
        return pd.NaT


source = "jp_eqhist_data.xlsx"
df = pd.read_excel(source)

df["Time"] = df["Time"].map(parse_japanese_datetime)

df["Magnitude"] = pd.to_numeric(df["Magnitude"], errors="coerce").fillna(0)

df.to_excel("jp_eqhist_data_filtered.xlsx", index=False, sheet_name="Data")
