# vietnam_fuel.py
import os
import re
import json
import datetime
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

BASE = os.path.dirname(os.path.abspath(__file__))
CSV_FILE = os.path.join(BASE, "越南历史油价.csv")
OUT_JSON = os.path.join(BASE, "vietnam_fuel.json")

URL = "https://www.petrolimex.com.vn/index.html"

# 品种：越南文 key -> 中文显示名 + CSV 前缀
FUEL_MAP = {
    "Xăng E10 RON 95 Mức 5": ("汽油 E10 RON 95（5 级）", "E10RON95_M5"),
    "Xăng E10 RON 95 Mức 3": ("汽油 E10 RON 95（3 级）", "E10RON95_M3"),
    "Xăng E5 RON 92 Mức 2":  ("汽油 E5 RON 92（2 级）",  "E5RON92_M2"),
    "DO 0,001S Mức 5":       ("柴油 DO 0.001S（5 级）",   "DO001S_M5"),
    "DO 0,05S Mức 2":        ("柴油 DO 0.05S（2 级）",    "DO005S_M2"),
    "Dầu hỏa 2-K":           ("煤油 2-K",                "Kero2K"),
}

# CSV 列
def make_csv_columns():
    cols = ["Date"]
    for _, (_, prefix) in FUEL_MAP.items():
        cols.append(f"{prefix}_Z1")
        cols.append(f"{prefix}_Z2")
        cols.append(f"{prefix}_Avg")
    return cols

CSV_COLUMNS = make_csv_columns()


def make_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    )
    return webdriver.Chrome(options=options)


def parse_price(s):
    """'29.050' -> 29050（越南盾）"""
    if not s:
        return None
    s = str(s).strip().replace(".", "").replace(",", "")
    m = re.search(r"\d+", s)
    return int(m.group()) if m else None


def fetch_prices():
    """用 Selenium 抓价格。返回 {越南品种名: {Z1: int, Z2: int}}"""
    driver = make_driver()
    try:
        driver.get(URL)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "table"))
        )
        time.sleep(8)  # 等 JS 填完

        # 用 outerHTML 抓（text 拿不到）
        tables = driver.find_elements(By.TAG_NAME, "table")
        if not tables:
            raise ValueError("没找到表格")

        # 直接解析 outerHTML
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(tables[0].get_attribute("outerHTML"), "lxml")

        result = {}
        for tr in soup.select("tbody tr"):
            tds = tr.find_all("td")
            if len(tds) < 3:
                continue
            name = tds[0].get_text(strip=True)
            z1 = parse_price(tds[1].get_text(strip=True))
            z2 = parse_price(tds[2].get_text(strip=True))
            result[name] = {"Z1": z1, "Z2": z2}
        return result
    finally:
        driver.quit()


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


def round2(v):
    return round(v * 100) / 100


def calc_changes(df, today_str, today_prices):
    """
    today_prices: {prefix: {Z1, Z2, Avg}}
    返回：{prefix: {prev, year_start, last_year}} + {"__AVERAGE__": 全品种平均}
    """
    if df.empty:
        return {}

    df = df.copy()
    df["_date"] = df["Date"].apply(parse_csv_date)
    df = df.dropna(subset=["_date"]).sort_values("_date").reset_index(drop=True)

    today_date = parse_csv_date(today_str)
    if not today_date:
        return {}

    result = {}

    # 每个品种：用 Avg 列算波动率
    for vn_name, (cn_name, prefix) in FUEL_MAP.items():
        col = f"{prefix}_Avg"
        today_val = today_prices.get(prefix, {}).get("Avg")
        if today_val is None or col not in df.columns:
            result[prefix] = {"prev": None, "year_start": None, "last_year": None}
            continue

        r = {"prev": None, "year_start": None, "last_year": None}

        # 较上期：找不同值的最近一天
        prev_rows = df[df["_date"] < today_date]
        for _, row in prev_rows.iloc[::-1].iterrows():
            v = row[col]
            try:
                v = float(v) if v not in (None, "", "nan") else None
            except ValueError:
                v = None
            if v is not None and abs(v - today_val) > 0.5:
                r["prev"] = round2((today_val - v) / v * 100)
                break

        # 较今年 1 月
        year = today_date.year
        ys_rows = df[(df["_date"].apply(lambda d: d.year) == year) &
                     (df["_date"].apply(lambda d: d.month) == 1)]
        if not ys_rows.empty:
            try:
                v = float(ys_rows.iloc[0][col])
                if v:
                    r["year_start"] = round2((today_val - v) / v * 100)
            except (ValueError, TypeError):
                pass

        # 较去年同期
        ly_date = today_date.replace(year=today_date.year - 1)
        ly_candidates = df[df["_date"].apply(lambda d: d.year) == (today_date.year - 1)].copy()
        if not ly_candidates.empty:
            ly_candidates["_diff"] = ly_candidates["_date"].apply(lambda d: abs((d - ly_date).days))
            ly_row = ly_candidates.sort_values("_diff").iloc[0]
            try:
                v = float(ly_row[col])
                if v:
                    r["last_year"] = round2((today_val - v) / v * 100)
            except (ValueError, TypeError):
                pass

        result[prefix] = r

    # 全品种平均（所有 Avg 的平均）
    today_avgs = [v["Avg"] for v in today_prices.values() if v.get("Avg") is not None]
    today_overall_avg = sum(today_avgs) / len(today_avgs) if today_avgs else None

    if today_overall_avg is not None:
        # 计算历史每天的"全品种平均"
        avg_cols = [f"{p}_Avg" for (_, (_, p)) in FUEL_MAP.items()]
        df["_overall_avg"] = df[avg_cols].apply(
            lambda row: _safe_avg(row), axis=1
        )

        r = {"prev": None, "year_start": None, "last_year": None}

        # 较上期
        prev_rows = df[df["_date"] < today_date]
        for _, row in prev_rows.iloc[::-1].iterrows():
            v = row["_overall_avg"]
            if v is not None and abs(v - today_overall_avg) > 0.5:
                r["prev"] = round2((today_overall_avg - v) / v * 100)
                break

        # 较今年 1 月
        year = today_date.year
        ys_rows = df[(df["_date"].apply(lambda d: d.year) == year) &
                     (df["_date"].apply(lambda d: d.month) == 1)]
        if not ys_rows.empty and ys_rows.iloc[0]["_overall_avg"]:
            v = ys_rows.iloc[0]["_overall_avg"]
            r["year_start"] = round2((today_overall_avg - v) / v * 100)

        # 较去年同期
        ly_date = today_date.replace(year=today_date.year - 1)
        ly_candidates = df[df["_date"].apply(lambda d: d.year) == (today_date.year - 1)].copy()
        if not ly_candidates.empty:
            ly_candidates["_diff"] = ly_candidates["_date"].apply(lambda d: abs((d - ly_date).days))
            ly_row = ly_candidates.dropna(subset=["_overall_avg"]).sort_values("_diff")
            if not ly_row.empty and ly_row.iloc[0]["_overall_avg"]:
                v = ly_row.iloc[0]["_overall_avg"]
                r["last_year"] = round2((today_overall_avg - v) / v * 100)

        result["__AVERAGE__"] = r

    return result


def _safe_avg(row):
    vals = []
    for v in row:
        try:
            f = float(v)
            if f > 0:
                vals.append(f)
        except (ValueError, TypeError):
            continue
    return sum(vals) / len(vals) if vals else None


def main():
    print(f"[越南] 抓取 {URL}")

    try:
        raw = fetch_prices()
        print(f"  抓到 {len(raw)} 个品种")
    except Exception as e:
        print(f"  抓取失败：{e}")
        return

    # 整理成 {prefix: {Z1, Z2, Avg}}
    today_prices = {}
    for vn_name, (cn_name, prefix) in FUEL_MAP.items():
        if vn_name not in raw:
            print(f"  [警告] 没抓到：{vn_name}")
            continue
        z1 = raw[vn_name].get("Z1")
        z2 = raw[vn_name].get("Z2")
        avg = round2((z1 + z2) / 2) if (z1 and z2) else None
        today_prices[prefix] = {"Z1": z1, "Z2": z2, "Avg": avg}
        print(f"  {cn_name}: Z1={z1} Z2={z2} Avg={avg}")

    # 日期（今天）
    today_str = datetime.datetime.now().strftime("%Y/%m/%d")

    # 构造新行
    new_row = {"Date": today_str}
    for prefix, vals in today_prices.items():
        new_row[f"{prefix}_Z1"] = vals["Z1"]
        new_row[f"{prefix}_Z2"] = vals["Z2"]
        new_row[f"{prefix}_Avg"] = vals["Avg"]

    # 更新 CSV
    df = load_csv()
    existing = df[df["Date"].astype(str).str.strip() == today_str]
    if not existing.empty:
        idx = existing.index[0]
        # 检查是否有变化
        same = True
        for col, val in new_row.items():
            if col == "Date":
                continue
            old_val = df.at[idx, col]
            try:
                old_f = float(old_val) if old_val not in (None, "", "nan") else None
            except ValueError:
                old_f = None
            new_f = float(val) if val is not None else None
            if old_f != new_f:
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

    # 算波动率
    df = load_csv()
    changes = calc_changes(df, today_str, today_prices)

    # 存 JSON
    result = {
        "update_time": (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S"),
        "date": today_str,
        "prices": today_prices,   # {prefix: {Z1, Z2, Avg}}
        "changes": changes,       # {prefix: {prev, year_start, last_year}}
        "fuel_names": {           # prefix -> 中文名
            prefix: cn_name for (_, (cn_name, prefix)) in FUEL_MAP.items()
        },
    }

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n已写入 {OUT_JSON}")


if __name__ == "__main__":
    main()