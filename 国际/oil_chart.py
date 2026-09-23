# oil_chart.py
import requests
import json
import re
import os
import time
import datetime
import pandas as pd

OUT_FILE = "oil_minline.json"
HISTORY_XLSX = "oil_history.xlsx"

SYMBOLS = {
    "CL":  "WTI美国",
    "OIL": "布伦特",
}

URL_TPL = ("https://stock2.finance.sina.com.cn/futures/api/jsonp.php/"
           "var%20_/GlobalFuturesService.getGlobalFuturesMinLine?symbol={}")

# ---------- Sheet1 列名 ----------
COL_DATE = "日期"
COL_WTI = "WTI纽约原油"
COL_BRENT = "布伦特油"
COL_OPEC = "OPEC一揽子"
SHEET1_COLS = [COL_DATE, COL_WTI, COL_BRENT, COL_OPEC]

# ---------- Sheet2 列名 ----------
COL2_DATE = "日期"
COL2_NAME = "品种"
COL2_OPEN = "开盘"
COL2_CLOSE = "收盘"
COL2_HIGH = "最高"
COL2_LOW = "最低"
SHEET2_COLS = [COL2_DATE, COL2_NAME, COL2_OPEN, COL2_CLOSE, COL2_HIGH, COL2_LOW]

# 新浪名 -> Sheet2 品种名
HISTORY_NAME_MAP = {
    "WTI美国": COL_WTI,
    "布伦特":  COL_BRENT,
}


# ---------- 抓分时 ----------
def fetch_minline(symbol):
    url = URL_TPL.format(symbol)
    r = requests.get(url, headers={
        "Referer": "https://finance.sina.com.cn",
        "User-Agent": "Mozilla/5.0"
    }, timeout=15)
    text = r.text

    m = re.search(r'var\s+_\s*\((\{.*\})\)\s*;?\s*$', text, re.S)
    if not m:
        m = re.search(r'var\s+_\s*=\s*\(?(\{.*\})\)?\s*;?\s*$', text, re.S)
    if not m:
        raise ValueError("解析 JSONP 失败：" + text[:200])

    obj = json.loads(m.group(1))
    rows = obj.get("minLine_1d", [])

    points = []
    open_price = None
    for idx, row in enumerate(rows):
        try:
            if len(row) >= 10:
                t = row[4]
                p = float(row[5])
                if idx == 0:
                    open_price = float(row[1])
            else:
                t = row[0]
                p = float(row[1])
            points.append([t, p])
        except (ValueError, IndexError):
            continue

    return {"points": points, "open": open_price}


# ---------- 交易日切分 ----------
def split_by_trading_day(points, ref_date):
    """
    输入：points = [[HH:MM, price], ...]
          ref_date = datetime.date，通常是"今天"
    输出：{trade_day: [points]}，trade_day 是 datetime.date
    交易日定义：06:00 ~ 次日 05:00
    """
    groups = {}
    cur_day = None

    for t, p in points:
        hh = int(t[:2])

        if hh < 6:
            # 属于"前一天"的交易日
            if cur_day is None:
                # 还没遇到 06:00 的点，往前推一天
                cur_day = ref_date - datetime.timedelta(days=1)
        else:
            # >= 06:00 → 属于"当天"的交易日
            if cur_day is None or hh >= 6:
                # 只有第一次遇到 >= 06:00，或者跨过 06:00 后，才切
                candidate = ref_date
                if cur_day is not None and (candidate - cur_day).days >= 1:
                    cur_day = candidate
                elif cur_day is None:
                    cur_day = candidate

        groups.setdefault(cur_day, []).append([t, p])

    return groups


def aggregate_day(points):
    if not points:
        return None
    prices = [p for _, p in points]
    return {
        "open":   prices[0],
        "close":  prices[-1],
        "high":   max(prices),
        "low":    min(prices),
        "count":  len(prices),
    }


# ---------- Sheet1 读写 ----------
def load_sheet1():
    if os.path.exists(HISTORY_XLSX):
        try:
            df = pd.read_excel(HISTORY_XLSX, sheet_name="Sheet1", dtype=object)
            df.columns = [str(c).strip() for c in df.columns]
        except Exception:
            df = pd.DataFrame(columns=SHEET1_COLS)
    else:
        df = pd.DataFrame(columns=SHEET1_COLS)

    for c in SHEET1_COLS:
        if c not in df.columns:
            df[c] = None
        df[c] = df[c].astype(object)

    df = df[SHEET1_COLS].copy()
    return df


def load_sheet2():
    if os.path.exists(HISTORY_XLSX):
        try:
            xls = pd.ExcelFile(HISTORY_XLSX)
            if "Sheet2" in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name="Sheet2", dtype=object)
                df.columns = [str(c).strip() for c in df.columns]
            else:
                df = pd.DataFrame(columns=SHEET2_COLS)
        except Exception:
            df = pd.DataFrame(columns=SHEET2_COLS)
    else:
        df = pd.DataFrame(columns=SHEET2_COLS)

    for c in SHEET2_COLS:
        if c not in df.columns:
            df[c] = None
        df[c] = df[c].astype(object)

    df = df[SHEET2_COLS].copy()
    return df


def save_both_sheets(df1, df2):
    with pd.ExcelWriter(HISTORY_XLSX, engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name="Sheet1", index=False)
        df2.to_excel(writer, sheet_name="Sheet2", index=False)


# ---------- 更新 Sheet1（只写收盘价） ----------
def upsert_sheet1(df1, trade_day_str, price_map):
    df1[COL_DATE] = df1[COL_DATE].apply(
        lambda x: str(x).strip() if x is not None and str(x).strip() else ""
    )
    mask = df1[COL_DATE] == trade_day_str
    if mask.any():
        idx = df1.index[mask][0]
        for col, val in price_map.items():
            if val is not None:
                df1.at[idx, col] = val
        print(f"[Sheet1] 更新 {trade_day_str}")
    else:
        new_row = {COL_DATE: trade_day_str, COL_WTI: None, COL_BRENT: None, COL_OPEC: None}
        new_row.update(price_map)
        df1 = pd.concat([df1, pd.DataFrame([new_row])], ignore_index=True)
        print(f"[Sheet1] 追加 {trade_day_str}")

    df1["_d"] = pd.to_datetime(df1[COL_DATE], errors="coerce")
    df1 = df1.sort_values("_d", na_position="last").drop(columns=["_d"])
    df1 = df1.reset_index(drop=True)
    return df1


# ---------- 更新 Sheet2（开/收/高/低） ----------
def upsert_sheet2(df2, trade_day_str, day_data_map):
    """
    day_data_map: {'WTI美国': {'open':..., 'close':..., 'high':..., 'low':...}, '布伦特': {...}}
    """
    df2[COL2_DATE] = df2[COL2_DATE].apply(
        lambda x: str(x).strip() if x is not None and str(x).strip() else ""
    )
    df2[COL2_NAME] = df2[COL2_NAME].apply(
        lambda x: str(x).strip() if x is not None and str(x).strip() else ""
    )

    for sym_name, aggr in day_data_map.items():
        cn = HISTORY_NAME_MAP.get(sym_name)
        if not cn or not aggr:
            continue

        mask = (df2[COL2_DATE] == trade_day_str) & (df2[COL2_NAME] == cn)
        if mask.any():
            idx = df2.index[mask][0]
            df2.at[idx, COL2_OPEN]  = aggr["open"]
            df2.at[idx, COL2_CLOSE] = aggr["close"]
            df2.at[idx, COL2_HIGH]  = aggr["high"]
            df2.at[idx, COL2_LOW]   = aggr["low"]
            print(f"[Sheet2] 更新 {trade_day_str} / {cn}")
        else:
            new_row = {
                COL2_DATE: trade_day_str, COL2_NAME: cn,
                COL2_OPEN: aggr["open"], COL2_CLOSE: aggr["close"],
                COL2_HIGH: aggr["high"], COL2_LOW: aggr["low"],
            }
            df2 = pd.concat([df2, pd.DataFrame([new_row])], ignore_index=True)
            print(f"[Sheet2] 追加 {trade_day_str} / {cn}")

    df2["_d"] = pd.to_datetime(df2[COL2_DATE], errors="coerce")
    df2 = df2.sort_values(["_d", COL2_NAME]).drop(columns=["_d"]).reset_index(drop=True)
    return df2


# ---------- 主流程 ----------
def main():
    result = {
        "update_time": (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S"),
        "series": {},
        "meta": {},
    }

    all_minlines = {}

    for sym, cn in SYMBOLS.items():
        try:
            d = fetch_minline(sym)
            result["series"][cn] = d["points"]
            result["meta"][cn] = {"open": d["open"]}
            all_minlines[cn] = d["points"]

            print(f"{cn}: {len(d['points'])} 个点，今开 {d['open']}，"
                  f"区间 {d['points'][0][0]} ~ {d['points'][-1][0]}")
        except Exception as e:
            print(f"{cn} 抓取失败：{e}")
            result["series"][cn] = []
            result["meta"][cn] = {}

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False)
    print(f"已写入 {OUT_FILE}")

    # ---------- 日聚合 ----------
    try:
        today = datetime.date.today()

        day_groups_wti = split_by_trading_day(all_minlines.get("WTI美国", []), today)
        day_groups_brent = split_by_trading_day(all_minlines.get("布伦特", []), today)

        all_days = set(day_groups_wti.keys()) | set(day_groups_brent.keys())

        df1 = load_sheet1()
        df2 = load_sheet2()

        for day in sorted(all_days):
            day_data_map = {}
            price_map = {}

            if day in day_groups_wti:
                aggr = aggregate_day(day_groups_wti[day])
                if aggr:
                    day_data_map["WTI美国"] = aggr
                    price_map[COL_WTI] = aggr["close"]
            if day in day_groups_brent:
                aggr = aggregate_day(day_groups_brent[day])
                if aggr:
                    day_data_map["布伦特"] = aggr
                    price_map[COL_BRENT] = aggr["close"]

            if day_data_map:
                day_str = day.strftime("%Y-%m-%d")
                df1 = upsert_sheet1(df1, day_str, price_map)
                df2 = upsert_sheet2(df2, day_str, day_data_map)

        save_both_sheets(df1, df2)
        print(f"[history] 已写回 {HISTORY_XLSX}（Sheet1 {len(df1)} 行，Sheet2 {len(df2)} 行）")
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[日聚合] 失败：{e}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3 and sys.argv[1] == "loop":
        interval = int(sys.argv[2])
        print(f"每 {interval} 秒刷新一次，Ctrl+C 停止")
        while True:
            main()
            time.sleep(interval)
    else:
        main()
