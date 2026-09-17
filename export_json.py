import pandas as pd
import json
import re
from datetime import date as _date

import os
BASE = os.path.dirname(os.path.abspath(__file__))
HISTORY_FILE  = os.path.join(BASE, "各省历史油价数据.xlsx")
CALENDAR_FILE = os.path.join(BASE, "油价调整日历.xlsx")
OUTPUT_JSON   = os.path.join(BASE, "data.json")
HTML_FILE     = os.path.join(BASE, "index.html")

# 与 测试.py 保持一致：31 个省级行政区
PROVINCES = [
    "北京市","天津市","河北省","山西省","辽宁省","吉林省","黑龙江省",
    "上海市","江苏省","浙江省","安徽省","福建省","江西省","山东省",
    "湖北省","湖南省","河南省","海南省","重庆市","广东省","广西壮族自治区",
    "宁夏","甘肃省","新疆",
    "内蒙古","四川省","贵州省","云南省","陕西省","青海省","西藏",
]

# ★ 17 个独立定价市（sheet 名 = "省-市"）
CITY_SHEETS = [
    "广东省-深圳市","辽宁省-大连市","山东省-青岛市","福建省-厦门市",
    "浙江省-宁波市","西藏-拉萨市","新疆-乌鲁木齐市","新疆-克拉玛依市",
    "内蒙古-呼和浩特市","四川省-甘孜藏族自治州","四川省-阿坝藏族羌族自治州",
    "四川省-凉山彝族自治州","云南省-迪庆藏族自治州","云南省-怒江傈僳族自治州",
    "青海省-玉树藏族自治州","青海省-果洛藏族自治州","甘肃省-甘南藏族自治州",
]

FUEL_COLS = ["89号汽油","92号汽油","95号汽油","98号汽油","0号柴油"]
NULL_TOKENS = {"", "-", "—", "无", "None", "nan", "NaN", "null"}


def to_float_or_none(v):
    if v is None or pd.isna(v):
        return None
    s = str(v).strip()
    if s in NULL_TOKENS:
        return None
    try:
        f = float(v)
    except (ValueError, TypeError):
        return None
    if f == 0:          # ★ 0 也视为无数据
        return None
    return f


def get_next_adjust():
    try:
        df = pd.read_excel(CALENDAR_FILE, header=None)
    except Exception as e:
        print("⚠ 读取日历表失败：", e)
        return None, None

    today = _date.today()
    future = []
    for v in df.iloc[1:, 0].dropna():
        s = str(v).strip()
        if not s:
            continue
        d = None
        for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d"):
            try:
                d = pd.to_datetime(s, format=fmt).date()
                break
            except Exception:
                continue
        if d is None:
            try:
                d = pd.to_datetime(s).date()
            except Exception:
                continue
        if d >= today:
            future.append(d)

    if not future:
        return None, None
    nxt = min(future)
    return nxt.strftime("%Y-%m-%d"), (nxt - today).days


def export():
    result = {}
    # ★ 省级 + 17 个独立定价市
    for p in PROVINCES + CITY_SHEETS:
        try:
            df = pd.read_excel(HISTORY_FILE, sheet_name=p)
        except Exception:
            continue
        df["日期"] = pd.to_datetime(df["日期"], errors="coerce").dt.strftime("%Y-%m-%d")
        df = df.dropna(subset=["日期"]).sort_values("日期", ascending=False).reset_index(drop=True)

        records = []
        for i, row in df.iterrows():
            rec = {"date": row["日期"]}
            for c in FUEL_COLS:
                cur = to_float_or_none(row.get(c))
                rec[c] = cur
                diff = None
                if cur is not None and i + 1 < len(df):
                    prev = to_float_or_none(df.iloc[i + 1].get(c))
                    if prev is not None:
                        diff = round(cur - prev, 2)
                rec[c + "_diff"] = diff
            records.append(rec)
        result[p] = records

    next_date, days_left = get_next_adjust()
    result["_meta"] = {
        "next_adjust": next_date,
        "days_left": days_left,
        "today": _date.today().strftime("%Y-%m-%d"),
    }

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"✅ 已导出 {len(result)-1} 个 sheet → {OUTPUT_JSON}")
    print(f"   下次调价：{next_date}（剩 {days_left} 天）")

    inject_into_html(result)


def inject_into_html(data):
    with open(HTML_FILE, "r", encoding="utf-8") as f:
        html = f.read()
    data_js = ('<script id="data-json" type="application/json">' +
               json.dumps(data, ensure_ascii=False) + '</script>')
    html = re.sub(r'<script id="data-json".*?</script>', '', html, flags=re.S)
    html = html.replace('</body>', data_js + '\n</body>')
    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ 数据已注入 index.html")


if __name__ == "__main__":
    export()