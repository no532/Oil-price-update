# oil_monthly.py
import json
import datetime
import os
import pandas as pd

# ★ 用脚本所在目录作为基准（不管在哪里跑都对）
BASE = os.path.dirname(os.path.abspath(__file__))
HISTORY_XLSX = os.path.join(BASE, "oil_history.xlsx")
OUT_JSON = os.path.join(BASE, "oil_monthly.json")

COL_DATE = "日期"
COL_WTI = "WTI纽约原油"
COL_BRENT = "布伦特油"
COL_OPEC = "OPEC一揽子"

COL2_DATE = "日期"
COL2_NAME = "品种"
COL2_OPEN = "开盘"
COL2_CLOSE = "收盘"
COL2_HIGH = "最高"
COL2_LOW = "最低"


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


def load_sheet1():
    """读 Sheet1：日期 / WTI / 布伦特 / OPEC"""
    if not os.path.exists(HISTORY_XLSX):
        print(f"未找到 {HISTORY_XLSX}")
        return pd.DataFrame(columns=[COL_DATE, COL_WTI, COL_BRENT, COL_OPEC])

    df = pd.read_excel(HISTORY_XLSX, sheet_name="Sheet1")
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


def load_sheet2():
    """读 Sheet2：日期 / 品种 / 开 / 收 / 高 / 低"""
    empty = pd.DataFrame(columns=[COL2_DATE, COL2_NAME, COL2_OPEN, COL2_CLOSE, COL2_HIGH, COL2_LOW])
    if not os.path.exists(HISTORY_XLSX):
        return empty

    try:
        xls = pd.ExcelFile(HISTORY_XLSX)
        if "Sheet2" not in xls.sheet_names:
            return empty
        df = pd.read_excel(xls, sheet_name="Sheet2")
    except Exception:
        return empty

    df.columns = [str(c).strip() for c in df.columns]
    for c in [COL2_DATE, COL2_NAME, COL2_OPEN, COL2_CLOSE, COL2_HIGH, COL2_LOW]:
        if c not in df.columns:
            df[c] = None
    df = df[[COL2_DATE, COL2_NAME, COL2_OPEN, COL2_CLOSE, COL2_HIGH, COL2_LOW]].copy()
    df[COL2_DATE] = pd.to_datetime(df[COL2_DATE], errors="coerce").dt.strftime("%Y-%m-%d")
    for c in [COL2_OPEN, COL2_CLOSE, COL2_HIGH, COL2_LOW]:
        df[c] = df[c].apply(parse_price)
    df = df.dropna(subset=[COL2_DATE, COL2_NAME]).reset_index(drop=True)
    return df


def export_json(df1, df2):
    result = {
        "update_time": (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S"),
        "series": {},
        "dailyDetail": {},
    }

    # ---------- series（Sheet1） ----------
    col_map = {
        COL_WTI: "WTI纽约原油",
        COL_BRENT: "布伦特油",
        COL_OPEC: "OPEC一揽子",
    }
    for col, name in col_map.items():
        pts = []
        for _, row in df1.iterrows():
            v = row[col]
            if v is not None and not pd.isna(v):
                pts.append([row[COL_DATE], float(v)])
        result["series"][name] = pts

    # ---------- dailyDetail（Sheet2） ----------
    detail = {}
    for _, row in df2.iterrows():
        name = row[COL2_NAME]
        date = row[COL2_DATE]
        if name not in detail:
            detail[name] = {}
        detail[name][date] = {
            "open":  float(row[COL2_OPEN])  if row[COL2_OPEN]  is not None and not pd.isna(row[COL2_OPEN])  else None,
            "close": float(row[COL2_CLOSE]) if row[COL2_CLOSE] is not None and not pd.isna(row[COL2_CLOSE]) else None,
            "high":  float(row[COL2_HIGH])  if row[COL2_HIGH]  is not None and not pd.isna(row[COL2_HIGH])  else None,
            "low":   float(row[COL2_LOW])   if row[COL2_LOW]   is not None and not pd.isna(row[COL2_LOW])   else None,
        }
    result["dailyDetail"] = detail

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False)
    print(f"已写入 {OUT_JSON}（Sheet1 {len(df1)} 行，Sheet2 {len(df2)} 行）")


def main():
    df1 = load_sheet1()
    df2 = load_sheet2()
    export_json(df1, df2)


if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3 and sys.argv[1] == "loop":
        interval = int(sys.argv[2])
        print(f"每 {interval} 秒刷新一次，Ctrl+C 停止")
        while True:
            main()
            import time
            time.sleep(interval)
    else:
        main()
