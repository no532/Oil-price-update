# oil_monthly.py
import json
import datetime
import os
import re
import pandas as pd

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
    """把 '90.043' / 'US $ 90.043' / '90.043美元/桶' / NaN 统一转成 float"""
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    s = s.replace("US", "").replace("$", "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        m = re.search(r"-?\d+(?:\.\d+)?", s)
        return float(m.group()) if m else None


def parse_date(v):
    """
    兼容多种日期格式：
      - '2026-09-22'
      - '2026-09-21 00:00:00'
      - datetime 对象
      - '2026/09/22'
    """
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    # 已经是 datetime 对象
    if hasattr(v, "strftime"):
        try:
            return v.strftime("%Y-%m-%d")
        except Exception:
            return None
    s = str(v).strip()
    if not s or s.lower() == "nat" or s.lower() == "nan":
        return None
    # 快速路径：YYYY-MM-DD 或 YYYY/MM/DD
    if len(s) >= 10 and s[4] in ("-", "/") and s[7] in ("-", "/"):
        return s[:10].replace("/", "-")
    # 兜底：pd.to_datetime
    try:
        d = pd.to_datetime(s, errors="coerce")
        if pd.isna(d):
            return None
        return d.strftime("%Y-%m-%d")
    except Exception:
        return None


def load_sheet1():
    if not os.path.exists(HISTORY_XLSX):
        print(f"未找到 {HISTORY_XLSX}")
        return pd.DataFrame(columns=[COL_DATE, COL_WTI, COL_BRENT, COL_OPEC])

    df = pd.read_excel(HISTORY_XLSX, sheet_name="Sheet1", dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    for c in [COL_DATE, COL_WTI, COL_BRENT, COL_OPEC]:
        if c not in df.columns:
            df[c] = None
    df = df[[COL_DATE, COL_WTI, COL_BRENT, COL_OPEC]].copy()
    df[COL_DATE] = df[COL_DATE].apply(parse_date)
    df[COL_WTI] = df[COL_WTI].apply(parse_price)
    df[COL_BRENT] = df[COL_BRENT].apply(parse_price)
    df[COL_OPEC] = df[COL_OPEC].apply(parse_price)
    df = df.dropna(subset=[COL_DATE]).sort_values(COL_DATE).reset_index(drop=True)
    return df


def load_sheet2():
    empty = pd.DataFrame(columns=[COL2_DATE, COL2_NAME, COL2_OPEN, COL2_CLOSE, COL2_HIGH, COL2_LOW])
    if not os.path.exists(HISTORY_XLSX):
        return empty

    try:
        xls = pd.ExcelFile(HISTORY_XLSX)
        if "Sheet2" not in xls.sheet_names:
            return empty
        df = pd.read_excel(xls, sheet_name="Sheet2", dtype=object)
    except Exception:
        return empty

    df.columns = [str(c).strip() for c in df.columns]
    for c in [COL2_DATE, COL2_NAME, COL2_OPEN, COL2_CLOSE, COL2_HIGH, COL2_LOW]:
        if c not in df.columns:
            df[c] = None
    df = df[[COL2_DATE, COL2_NAME, COL2_OPEN, COL2_CLOSE, COL2_HIGH, COL2_LOW]].copy()
    df[COL2_DATE] = df[COL2_DATE].apply(parse_date)
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
