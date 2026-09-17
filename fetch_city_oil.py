# fetch_city_oil.py
import requests
import re
import time
import os
import pandas as pd
from datetime import date as _date

BASE = os.path.dirname(os.path.abspath(__file__))
HISTORY_FILE = os.path.join(BASE, "各省历史油价数据.xlsx")
LAST_FETCH_FILE = os.path.join(BASE, "fetch_city_last.txt")

# 17 个独立定价市：sheet 名 → 本地市名 → 行政区划代码
CITIES = [
    ("广东省-深圳市",              "深圳市",           "440300"),
    ("辽宁省-大连市",              "大连市",           "210200"),
    ("山东省-青岛市",              "青岛市",           "370200"),
    ("福建省-厦门市",              "厦门市",           "350200"),
    ("浙江省-宁波市",              "宁波市",           "330200"),
    ("西藏-拉萨市",                "拉萨市",           "540100"),
    ("新疆-乌鲁木齐市",            "乌鲁木齐市",       "650100"),
    ("新疆-克拉玛依市",            "克拉玛依市",       "650200"),
    ("内蒙古-呼和浩特市",          "呼和浩特市",       "150100"),
    ("四川省-甘孜藏族自治州",      "甘孜藏族自治州",   "513300"),
    ("四川省-阿坝藏族羌族自治州",  "阿坝藏族羌族自治州","513200"),
    ("四川省-凉山彝族自治州",      "凉山彝族自治州",   "513400"),
    ("云南省-迪庆藏族自治州",      "迪庆藏族自治州",   "533400"),
    ("云南省-怒江傈僳族自治州",    "怒江傈僳族自治州", "533300"),
    ("青海省-玉树藏族自治州",      "玉树藏族自治州",   "632700"),
    ("青海省-果洛藏族自治州",      "果洛藏族自治州",   "632600"),
    ("甘肃省-甘南藏族自治州",      "甘南藏族自治州",   "623000"),
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/120.0 Safari/537.36"
}

FUEL_COLS = ["89号汽油","92号汽油","95号汽油","98号汽油","0号柴油"]


def already_fetched_today():
    today = _date.today().strftime("%Y-%m-%d")
    if os.path.exists(LAST_FETCH_FILE):
        with open(LAST_FETCH_FILE, "r", encoding="utf-8") as f:
            last = f.read().strip()
        if last == today:
            return True
    return False


def mark_fetched_today():
    today = _date.today().strftime("%Y-%m-%d")
    with open(LAST_FETCH_FILE, "w", encoding="utf-8") as f:
        f.write(today)


def fetch_city_history(code):
    url = f"https://www.icauto.com.cn/oil/price_{code}_0.html"
    r = requests.get(url, headers=HEADERS, timeout=20)
    r.encoding = "utf-8"
    html = r.text

    m = re.search(r"historyOil['\"]>(.*?)</table>", html, re.S)
    if not m:
        return []

    table_html = m.group(1)
    rows = re.findall(r"<tr>(.*?)</tr>", table_html, re.S)
    result = []
    for row in rows:
        tds = re.findall(r"<td>(.*?)</td>", row, re.S)
        if len(tds) < 7:
            continue
        date = tds[0].strip()
        if not re.match(r"\d{4}-\d{2}-\d{2}", date):
            continue
        vals = [t.strip() for t in tds[1:7]]
        result.append([date] + vals)
    return result


def to_float(v):
    try:
        f = float(v)
        return None if f == 0 else f
    except (ValueError, TypeError):
        return None


def build_records(raw_rows):
    records = []
    for r in raw_rows:
        date = r[0]
        v92 = to_float(r[2])
        v95 = to_float(r[4])
        v0  = to_float(r[6])
        records.append({
            "日期": date,
            "89号汽油": None,
            "92号汽油": v92,
            "95号汽油": v95,
            "98号汽油": None,
            "0号柴油":  v0,
        })
    return records


def merge_into_excel(sheet_name, new_records):
    try:
        df_old = pd.read_excel(HISTORY_FILE, sheet_name=sheet_name)
        df_old["日期"] = pd.to_datetime(df_old["日期"], errors="coerce").dt.strftime("%Y-%m-%d")
    except Exception:
        df_old = pd.DataFrame(columns=["日期"] + FUEL_COLS)

    df_new = pd.DataFrame(new_records)
    df_all = pd.concat([df_new, df_old], ignore_index=True)
    df_all = df_all.drop_duplicates(subset=["日期"], keep="last")   # 旧数据（手动填的）优先
    df_all = df_all.sort_values("日期", ascending=False).reset_index(drop=True)

    with pd.ExcelWriter(HISTORY_FILE, engine="openpyxl", mode="a",
                        if_sheet_exists="replace") as writer:
        df_all.to_excel(writer, sheet_name=sheet_name, index=False)

    return len(df_all), len(df_new)


def main(force=False):
    if not force and already_fetched_today():
        print("→ 今日已爬过市级数据，跳过")
        return

    print("→ 开始爬取 17 个独立定价市")
    for sheet_name, local_name, code in CITIES:
        try:
            raw = fetch_city_history(code)
            if not raw:
                print(f"  ⚠ {sheet_name}: 未获取到数据")
                continue
            recs = build_records(raw)
            total, added = merge_into_excel(sheet_name, recs)
            print(f"  ✅ {sheet_name}: 抓 {len(recs)} 条，合并后共 {total} 条")
        except Exception as e:
            print(f"  ❌ {sheet_name}: {e}")
        time.sleep(1)

    mark_fetched_today()
    print("→ 完成，已记录今日爬取时间")


if __name__ == "__main__":
    main()
