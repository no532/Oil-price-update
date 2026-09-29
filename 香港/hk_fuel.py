# hk_fuel.py
import os
import re
import json
import datetime
import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
CSV_FILE = os.path.join(BASE, "香港历史油价.csv")
OUT_JSON = os.path.join(BASE, "hk_fuel.json")

DIESEL_URL = "https://oil-price.consumer.org.hk/tc/diesel"
OPEN_DATA_URL = "https://www.consumer.org.hk/pricewatch/oilwatch/opendata/oilprice.json"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
    "Accept-Language": "zh-HK,zh;q=0.9",
}

CSV_COLUMNS = [
    "Date",
    "中石化 (零售牌價)", "中國石油 (零售牌價)", "加德士 (零售牌價)", "埃索 (零售牌價)", "蜆殼 (零售牌價)",
    "中石化 (扣除門市折扣及燃油稅)", "中國石油 (扣除門市折扣及燃油稅)", "加德士 (扣除門市折扣及燃油稅)",
    "埃索 (扣除門市折扣及燃油稅)", "蜆殼 (扣除門市折扣及燃油稅)",
]

BRAND_MAP = {
    "中石化":   "中石化",
    "中國石油": "中國石油",
    "蜆殼":     "蜆殼",
    "埃索":     "埃索",
    "加德士":   "加德士",
}

AFTER_DISCOUNT_COLS = {
    "中石化":   "中石化 (扣除門市折扣及燃油稅)",
    "中國石油": "中國石油 (扣除門市折扣及燃油稅)",
    "蜆殼":     "蜆殼 (扣除門市折扣及燃油稅)",
    "埃索":     "埃索 (扣除門市折扣及燃油稅)",
    "加德士":   "加德士 (扣除門市折扣及燃油稅)",
}

RETAIL_COLS = {
    "中石化":   "中石化 (零售牌價)",
    "中國石油": "中國石油 (零售牌價)",
    "蜆殼":     "蜆殼 (零售牌價)",
    "埃索":     "埃索 (零售牌價)",
    "加德士":   "加德士 (零售牌價)",
}


def parse_price(s):
    if not s:
        return None
    s = s.replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None


def fetch_diesel():
    r = requests.get(DIESEL_URL, headers=HEADERS, timeout=20)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "lxml")

    date_str = None
    date_el = soup.select_one(".title-date")
    if date_el:
        text = date_el.get_text(strip=True)
        m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", text)
        if m:
            date_str = f"{m.group(1)}/{int(m.group(2))}/{int(m.group(3))}"

    data = {}
    for row in soup.select("a.board__row"):
        cells = row.select(".board__cell")
        if len(cells) < 4:
            continue
        name_el = cells[0].select_one(".clogo__name")
        brand = name_el.get_text(strip=True) if name_el else None
        if not brand:
            continue
        retail = parse_price(cells[1].get_text())
        after = parse_price(cells[2].get_text())
        discount = None
        if retail is not None and after is not None:
            discount = round(retail - after, 2)
        data[brand] = {
            "retail": retail,
            "after_discount": after,
            "discount": discount,
        }

    return {"date": date_str, "data": data}


def fetch_open_data():
    r = requests.get(OPEN_DATA_URL, headers=HEADERS, timeout=20)
    r.raise_for_status()
    arr = r.json()

    result = {}
    for group in arr:
        type_name = group.get("type", {}).get("tc")
        if not type_name:
            continue
        prices = {}
        for item in group.get("prices", []):
            vendor = item.get("vendor", {}).get("tc")
            price = item.get("price")
            if vendor and price:
                prices[vendor] = float(price)
        result[type_name] = prices
    return result


def load_csv():
    if not os.path.exists(CSV_FILE):
        return pd.DataFrame(columns=CSV_COLUMNS)
    df = pd.read_csv(CSV_FILE, dtype=str, encoding="utf-8-sig")
    df.columns = [str(c).strip() for c in df.columns]
    for c in CSV_COLUMNS:
        if c not in df.columns:
            df[c] = None
    return df[CSV_COLUMNS].copy()


def save_csv(df):
    df.to_csv(CSV_FILE, index=False, encoding="utf-8-sig")


def parse_csv_date(s):
    if not s:
        return None
    s = str(s).strip()
    for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def calc_changes(df, today_str, diesel_today):
    if df.empty:
        return {}

    df = df.copy()
    df["_date"] = df["Date"].apply(parse_csv_date)
    df = df.dropna(subset=["_date"]).sort_values("_date").reset_index(drop=True)

    today_date = parse_csv_date(today_str)
    if not today_date:
        return {}

    def avg_of_row(row, cols):
        vals = []
        for col in cols:
            v = row[col]
            if v is not None and str(v).strip():
                p = parse_price(v)
                if p:
                    vals.append(p)
        return sum(vals) / len(vals) if vals else None

    df["_avg"] = df.apply(lambda r: avg_of_row(r, AFTER_DISCOUNT_COLS.values()), axis=1)
    df["_avg_r"] = df.apply(lambda r: avg_of_row(r, RETAIL_COLS.values()), axis=1)

    result = {}

    # ========== 折后价 ==========
    for brand, col in AFTER_DISCOUNT_COLS.items():
        today_price = diesel_today.get(brand, {}).get("after_discount")
        if today_price is None:
            continue

        result[brand] = {"prev": None, "year_start": None, "last_year": None}

        prev_rows = df[df["_date"] < today_date]
        for _, row in prev_rows.iloc[::-1].iterrows():
            v = parse_price(row[col])
            if v is not None and abs(v - today_price) > 0.001:
                result[brand]["prev"] = round((today_price - v) / v * 100, 2)
                break

        year = today_date.year
        ys_rows = df[(df["_date"].apply(lambda d: d.year) == year) &
                     (df["_date"].apply(lambda d: d.month) == 1)]
        if not ys_rows.empty:
            p = parse_price(ys_rows.iloc[0][col])
            if p:
                result[brand]["year_start"] = round((today_price - p) / p * 100, 2)

        ly_date = today_date.replace(year=today_date.year - 1)
        ly_candidates = df[df["_date"].apply(lambda d: d.year) == (today_date.year - 1)].copy()
        if not ly_candidates.empty:
            ly_candidates["_diff"] = ly_candidates["_date"].apply(lambda d: abs((d - ly_date).days))
            ly_row = ly_candidates.sort_values("_diff").iloc[0]
            p = parse_price(ly_row[col])
            if p:
                result[brand]["last_year"] = round((today_price - p) / p * 100, 2)

    # 平均折后价
    today_avg_vals = [d["after_discount"] for d in diesel_today.values() if d.get("after_discount")]
    today_avg = sum(today_avg_vals) / len(today_avg_vals) if today_avg_vals else None

    if today_avg is not None:
        avg_changes = {"prev": None, "year_start": None, "last_year": None}

        prev_rows = df[df["_date"] < today_date]
        for _, row in prev_rows.iloc[::-1].iterrows():
            if row["_avg"] and abs(row["_avg"] - today_avg) > 0.001:
                avg_changes["prev"] = round((today_avg - row["_avg"]) / row["_avg"] * 100, 2)
                break

        year = today_date.year
        ys_rows = df[(df["_date"].apply(lambda d: d.year) == year) &
                     (df["_date"].apply(lambda d: d.month) == 1)]
        if not ys_rows.empty and ys_rows.iloc[0]["_avg"]:
            p = ys_rows.iloc[0]["_avg"]
            avg_changes["year_start"] = round((today_avg - p) / p * 100, 2)

        ly_date = today_date.replace(year=today_date.year - 1)
        ly_candidates = df[df["_date"].apply(lambda d: d.year) == (today_date.year - 1)].copy()
        if not ly_candidates.empty:
            ly_candidates["_diff"] = ly_candidates["_date"].apply(lambda d: abs((d - ly_date).days))
            ly_row = ly_candidates.dropna(subset=["_avg"]).sort_values("_diff")
            if not ly_row.empty and ly_row.iloc[0]["_avg"]:
                p = ly_row.iloc[0]["_avg"]
                avg_changes["last_year"] = round((today_avg - p) / p * 100, 2)

        result["__AVERAGE__"] = avg_changes

    # ========== 零售价 ==========
    result_retail = {}
    for brand, col in RETAIL_COLS.items():
        today_price = diesel_today.get(brand, {}).get("retail")
        if today_price is None:
            continue
        result_retail[brand] = {"prev": None, "year_start": None, "last_year": None}

        prev_rows = df[df["_date"] < today_date]
        for _, row in prev_rows.iloc[::-1].iterrows():
            v = parse_price(row[col])
            if v is not None and abs(v - today_price) > 0.001:
                result_retail[brand]["prev"] = round((today_price - v) / v * 100, 2)
                break

        year = today_date.year
        ys_rows = df[(df["_date"].apply(lambda d: d.year) == year) &
                     (df["_date"].apply(lambda d: d.month) == 1)]
        if not ys_rows.empty:
            p = parse_price(ys_rows.iloc[0][col])
            if p:
                result_retail[brand]["year_start"] = round((today_price - p) / p * 100, 2)

        ly_date = today_date.replace(year=today_date.year - 1)
        ly_candidates = df[df["_date"].apply(lambda d: d.year) == (today_date.year - 1)].copy()
        if not ly_candidates.empty:
            ly_candidates["_diff"] = ly_candidates["_date"].apply(lambda d: abs((d - ly_date).days))
            ly_row = ly_candidates.sort_values("_diff").iloc[0]
            p = parse_price(ly_row[col])
            if p:
                result_retail[brand]["last_year"] = round((today_price - p) / p * 100, 2)

    # 平均零售价
    today_avg_vals_r = [d["retail"] for d in diesel_today.values() if d.get("retail")]
    today_avg_r = sum(today_avg_vals_r) / len(today_avg_vals_r) if today_avg_vals_r else None
    if today_avg_r is not None:
        avg_r = {"prev": None, "year_start": None, "last_year": None}
        prev_rows = df[df["_date"] < today_date]
        for _, row in prev_rows.iloc[::-1].iterrows():
            if row["_avg_r"] and abs(row["_avg_r"] - today_avg_r) > 0.001:
                avg_r["prev"] = round((today_avg_r - row["_avg_r"]) / row["_avg_r"] * 100, 2)
                break

        year = today_date.year
        ys_rows = df[(df["_date"].apply(lambda d: d.year) == year) &
                     (df["_date"].apply(lambda d: d.month) == 1)]
        if not ys_rows.empty and ys_rows.iloc[0]["_avg_r"]:
            p = ys_rows.iloc[0]["_avg_r"]
            avg_r["year_start"] = round((today_avg_r - p) / p * 100, 2)

        ly_candidates = df[df["_date"].apply(lambda d: d.year) == (today_date.year - 1)].copy()
        if not ly_candidates.empty:
            ly_candidates["_diff"] = ly_candidates["_date"].apply(lambda d: abs((d - ly_date).days))
            ly_row = ly_candidates.dropna(subset=["_avg_r"]).sort_values("_diff")
            if not ly_row.empty and ly_row.iloc[0]["_avg_r"]:
                p = ly_row.iloc[0]["_avg_r"]
                avg_r["last_year"] = round((today_avg_r - p) / p * 100, 2)

        result_retail["__AVERAGE__"] = avg_r

    result["__RETAIL__"] = result_retail

    return result


def main():
    try:
        diesel = fetch_diesel()
        print(f"[柴油] {diesel['date']} 抓到 {len(diesel['data'])} 家")
    except Exception as e:
        print(f"[柴油] 抓取失败：{e}")
        return

    if not diesel["date"]:
        print("没抓到日期，跳过")
        return

    open_data = {}
    try:
        open_data = fetch_open_data()
        print(f"\n[公开数据] 抓到 {len(open_data)} 种油品")
    except Exception as e:
        print(f"\n[公开数据] 抓取失败：{e}")

    today_str = diesel["date"]

    df = load_csv()
    new_row = {"Date": today_str}
    for brand, csv_prefix in BRAND_MAP.items():
        d = diesel["data"].get(brand, {})
        new_row[f"{csv_prefix} (零售牌價)"] = d.get("retail")
        new_row[f"{csv_prefix} (扣除門市折扣及燃油稅)"] = d.get("after_discount")

    existing = df[df["Date"].astype(str).str.strip() == today_str]
    if not existing.empty:
        idx = existing.index[0]
        same = True
        for col, val in new_row.items():
            if col == "Date":
                continue
            old_val = df.at[idx, col]
            old_val_f = float(old_val) if old_val not in (None, "", "nan") else None
            new_val_f = float(val) if val is not None else None
            if old_val_f != new_val_f:
                same = False
                break
        if same:
            print(f"\n[CSV] {today_str} 数据相同，跳过")
        else:
            for col, val in new_row.items():
                df.at[idx, col] = val
            save_csv(df)
            print(f"\n[CSV] {today_str} 已更新")
    else:
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_csv(df)
        print(f"\n[CSV] {today_str} 已追加")

    df = load_csv()
    changes = calc_changes(df, today_str, diesel["data"])

    avg_vals = [d["after_discount"] for d in diesel["data"].values() if d.get("after_discount")]
    avg_today = round(sum(avg_vals) / len(avg_vals), 2) if avg_vals else None

    avg_vals_r = [d["retail"] for d in diesel["data"].values() if d.get("retail")]
    avg_today_r = round(sum(avg_vals_r) / len(avg_vals_r), 2) if avg_vals_r else None

    result = {
        "update_time": (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S"),
        "date": today_str,
        "diesel": diesel["data"],
        "avg_today": avg_today,
        "avg_today_retail": avg_today_r,
        "changes": changes,
        "open_data": open_data,
    }

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n已写入 {OUT_JSON}")


if __name__ == "__main__":
    main()