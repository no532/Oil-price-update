import requests
import json
import re
import time
import datetime

OUT_FILE = "oil_minline.json"
SNAPSHOT_FILE = "oil_data.json"

SYMBOLS = {
    "CL":  "WTI美国",
    "OIL": "布伦特",
}

URL_TPL = "https://stock2.finance.sina.com.cn/futures/api/jsonp.php/var%20_/GlobalFuturesService.getGlobalFuturesMinLine?symbol={}"

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
                    open_price = float(row[1])   # 第一条的第2个字段 = 今开
            else:
                t = row[0]
                p = float(row[1])
            points.append([t, p])
        except (ValueError, IndexError):
            continue

    return {"points": points, "open": open_price}

def load_snapshot():
    """读 66laji 的快照，取昨收/最高/最低"""
    try:
        with open(SNAPSHOT_FILE, encoding="utf-8") as f:
            data = json.load(f)
        m = {}
        for it in data.get("items", []):
            m[it["name"]] = {
                "prev_close": it.get("prev_close"),
                "high": it.get("high"),
                "low": it.get("low"),
            }
        return m
    except Exception as e:
        print("读快照失败（不影响分时）：", e)
        return {}

def main():
    snap = load_snapshot()
    result = {
        "update_time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "series": {},
        "meta": {},
    }

    for sym, cn in SYMBOLS.items():
        try:
            d = fetch_minline(sym)
            result["series"][cn] = d["points"]

            # 补上今开 + 66laji 的昨收/最高/最低
            info = {"open": d["open"]}
            # 新浪 WTI 叫 "WTI美国"，66laji 叫 "WTI纽约原油"，映射一下
            snap_name = "WTI纽约原油" if cn == "WTI美国" else "布伦特原油"
            s = snap.get(snap_name, {})
            info["prev_close"] = s.get("prev_close")
            info["high"] = s.get("high")
            info["low"] = s.get("low")
            result["meta"][cn] = info

            print(f"{cn}: {len(d['points'])} 个点，今开 {d['open']}")
        except Exception as e:
            print(f"{cn} 抓取失败：{e}")
            result["series"][cn] = []
            result["meta"][cn] = {}

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False)

    print(f"已写入 {OUT_FILE}")

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
