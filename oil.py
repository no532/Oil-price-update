import requests
import csv
import json
import os
import time
import datetime

CSV_FILE = "oil.csv"
JSON_FILE = "oil_data.json"

API_URL = "https://api.66laji.cn/api/qqyoujia/index.php"
API_KEY = "253e9ba992e730afaa5a1a6e227f2bd7"

def fetch_all():
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    r = requests.get(API_URL, params={"apikey": API_KEY}, timeout=15)
    r.raise_for_status()
    j = r.json()

    if j.get("code") != 0:
        print(f"[{now}] 接口返回错误：{j.get('msg')}")
        return []

    rows = []
    for it in j["data"]["list"]:
        rows.append({
            "fetch_time": now,
            "name": it["name"],
            "price": it["latest_price"],
            "change": it.get("change"),
            "change_percent": it.get("change_percent"),
            "prev_close": it.get("yesterday_close"),
            "high": it.get("high"),
            "low": it.get("low"),
            "update_time": it.get("update_time"),
        })
    return rows

def save_csv(rows):
    exists = os.path.exists(CSV_FILE)
    with open(CSV_FILE, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["抓取时间","品种","最新价","涨跌","涨跌幅","昨收","最高","最低","更新时间"])
        for r in rows:
            w.writerow([r["fetch_time"], r["name"], r["price"], r["change"],
                        r["change_percent"], r["prev_close"], r["high"], r["low"],
                        r["update_time"]])

def save_json(rows):
    data = {
        "update_time": rows[0]["fetch_time"] if rows else "",
        "count": len(rows),
        "items": rows,
    }
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def run_once():
    rows = fetch_all()
    if not rows:
        return
    for r in rows:
        print(f"{r['name']}: {r['price']}  {r['change']}  {r['change_percent']}")
    save_csv(rows)
    save_json(rows)
    print(f"已更新 {JSON_FILE}（{len(rows)} 个品种）")

if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3 and sys.argv[1] == "loop":
        interval = int(sys.argv[2])
        print(f"每 {interval} 秒刷新一次，Ctrl+C 停止")
        while True:
            run_once()
            time.sleep(interval)
    else:
        run_once()