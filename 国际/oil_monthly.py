# oil_monthly.py
import json
import datetime
import os
import pandas as pd

HISTORY_XLSX = "oil_history.xlsx"
OUT_JSON = "oil_monthly.json"

COL_DATE = "日期"
COL_WTI = "WTI纽约原油"
COL_BRENT = "布伦特油"
COL_OPEC = "OPEC一揽子"


def parse_price(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    s = s.replace("US", "").replace("$", "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def load_history():
    if not os.path.exists(HISTORY_XLSX):
        print(f"未找到 {HISTORY_XLSX}")
        return pd.DataFrame(columns=[COL_DATE, COL_WTI, COL_BRENT, COL_OPEC])

    df = pd.read_excel(HISTORY_XLSX, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    for c in [COL_DATE, COL_WTI, COL_BRENT, COL_OPEC]:
        if c not in df.columns:
            df[c] = None
    df = df[[COL_DATE, COL_WTI, COL_BRENT, COL_OPEC]].copy()
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce").dt.strftime("%Y-%m-%d")
    df[COL_WTI] = df[COL_WTI].apply(parse_price)
    df[COL_BRENT] = df[COL_BRENT].apply(parse_price)
    df[COL_OPEC] = df[COL_OPEC].apply(parse_price)
    df = df.dropna(subset=[COL_DATE]).sort_values(COL_DATE).reset_index(drop=True)
    return df


def export_json(df):
    result = {
        "update_time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "series": {},
    }
    col_map = {
        COL_WTI: "WTI纽约原油",
        COL_BRENT: "布伦特油",
        COL_OPEC: "OPEC一揽子",
    }
    for col, name in col_map.items():
        pts = []
        for _, row in df.iterrows():
            v = row[col]
            if v is not None and not pd.isna(v):
                pts.append([row[COL_DATE], float(v)])
        result["series"][name] = pts

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False)
    print(f"已写入 {OUT_JSON}，共 {len(df)} 天")


if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3 and sys.argv[1] == "loop":
        interval = int(sys.argv[2])
        print(f"每 {interval} 秒刷新一次，Ctrl+C 停止")
        while True:
            df = load_history()
            export_json(df)
            import time
            time.sleep(interval)
    else:
        df = load_history()
        export_json(df)