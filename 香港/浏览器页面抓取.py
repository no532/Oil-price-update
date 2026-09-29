# hk_fuel.py
import os
import re
import json
import time
import datetime
import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
HISTORY_XLSX = os.path.join(BASE, "hk_fuel_history.xlsx")
OUT_JSON = os.path.join(BASE, "hk_fuel.json")

# 三种油品的 URL
URLS = {
    "无铅汽油":     "https://oil-price.consumer.org.hk/tc/",
    "特级无铅汽油": "https://oil-price.consumer.org.hk/tc/premium-petrol",
    "柴油":        "https://oil-price.consumer.org.hk/tc/diesel",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
    "Accept-Language": "zh-HK,zh;q=0.9",
}


def parse_price(s):
    """'$38.13' / '(-$13.0)' / '$25.13/升' -> float"""
    if not s:
        return None
    s = s.replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None


def fetch_one(url):
    """抓一个油品页面，返回 {date, rows}"""
    r = requests.get(url, headers=HEADERS, timeout=20)
    r.raise_for_status()
    html = r.text
    soup = BeautifulSoup(html, "lxml")

    # ---- 更新日期 ----
    date_str = None
    date_el = soup.select_one(".title-date")
    if date_el:
        text = date_el.get_text(strip=True)
        m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", text)
        if m:
            date_str = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"

    # ---- 每家品牌（board__row） ----
    rows = []
    for row in soup.select("a.board__row"):
        cells = row.select(".board__cell")
        if len(cells) < 4:
            continue
        # 第 1 格：品牌名
        name_el = cells[0].select_one(".clogo__name")
        brand = name_el.get_text(strip=True) if name_el else None
        if not brand:
            continue

        retail = parse_price(cells[1].get_text())
        after_discount = parse_price(cells[2].get_text())
        discount = parse_price(cells[3].get_text())

        rows.append({
            "brand": brand,
            "retail": retail,              # 零售牌价
            "after_discount": after_discount,  # 折后价（未税）
            "discount": discount,          # 门市折扣（负值）
        })

    return {"date": date_str, "rows": rows}


def main():
    result = {
        "update_time": (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S"),
        "date": None,
        "fuels": {},
    }

    for fuel_type, url in URLS.items():
        try:
            data = fetch_one(url)
            result["fuels"][fuel_type] = data
            if data["date"]:
                result["date"] = data["date"]
            print(f"[{fuel_type}] {data['date']} 抓到 {len(data['rows'])} 家")
            for r in data["rows"]:
                print(f"  {r['brand']}: 零售 {r['retail']} / 折后 {r['after_discount']} / 折扣 {r['discount']}")
        except Exception as e:
            print(f"[{fuel_type}] 失败：{e}")
            result["fuels"][fuel_type] = {"date": None, "rows": []}

    # 存 JSON
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n已写入 {OUT_JSON}")


if __name__ == "__main__":
    main()