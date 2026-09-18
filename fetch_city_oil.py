import requests
import re
import time
import os
import pandas as pd
from datetime import date as _date, datetime, timedelta

# ===================== 配置区 =====================
HISTORY_FILE = r"C:\Users\31007384\Desktop\新建文件夹\各省历史油价数据.xlsx"
LAST_FETCH_FILE = os.path.join(os.path.dirname(HISTORY_FILE), "last_fetch.txt")

# 城市列表：（Excel sheet 名，城市中文名，icauto 行政区划代码）
CITIES = [
    ("广东省-深圳市",              "深圳市",         "440300"),
    ("辽宁省-大连市",              "大连市",         "210200"),
    ("山东省-青岛市",              "青岛市",         "370200"),
    ("福建省-厦门市",              "厦门市",         "350200"),
    ("浙江省-宁波市",              "宁波市",         "330200"),
    ("西藏-拉萨市",                "拉萨市",         "540100"),
    ("新疆-乌鲁木齐市",            "乌鲁木齐市",     "650100"),
    ("新疆-克拉玛依市",            "克拉玛依市",     "650200"),
    ("内蒙古-呼和浩特市",          "呼和浩特市",     "150100"),
    ("四川省-甘孜藏族自治州",      "甘孜",           "513300"),
    ("四川省-阿坝藏族羌族自治州",  "阿坝",           "513200"),
    ("四川省-凉山彝族自治州",      "凉山",           "513400"),
    ("云南省-迪庆藏族自治州",      "迪庆",           "533400"),
    ("云南省-怒江傈僳族自治州",    "怒江",           "533300"),
    ("青海省-玉树藏族自治州",      "玉树",           "632700"),
    ("青海省-果洛藏族自治州",      "果洛",           "632600"),
    ("甘肃省-甘南藏族自治州",      "甘南",           "623000"),
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/120.0 Safari/537.36"
}

# 允许日期前后 ±2 天匹配
DAY_TOLERANCE = 2

# Excel 原始列名 → 代码内部列名
MAP_EXCEL_TO_INNER = {
    "调价日期": "日期",
    "日期":     "日期",
    "89汽油":   "89号汽油",
    "92汽油":   "92号汽油",
    "95汽油":   "95号汽油",
    "98汽油":   "98号汽油",
    "0柴油":    "0号柴油",
    "89号汽油": "89号汽油",
    "92号汽油": "92号汽油",
    "95号汽油": "95号汽油",
    "98号汽油": "98号汽油",
    "0号柴油":  "0号柴油",
}
OIL_COLS = ["89号汽油", "92号汽油", "95号汽油", "98号汽油", "0号柴油"]


# ===================== 工具函数 =====================
def parse_date_to_obj(d_val):
    """兼容 datetime、YYYY-MM-DD、YYYY/MM/DD、带时间"""
    if pd.isna(d_val):
        return None
    if isinstance(d_val, datetime):
        return d_val.date()
    if isinstance(d_val, _date):
        return d_val
    d_str = str(d_val).strip()
    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d", "%Y/%m/%d %H:%M:%S"]:
        try:
            return datetime.strptime(d_str, fmt).date()
        except ValueError:
            continue
    return None


def date_obj_to_str(d_obj):
    """date 对象 → YYYY/M/D 格式字符串（跟 Excel 里保持一致）"""
    if d_obj is None:
        return ""
    return f"{d_obj.year}/{d_obj.month}/{d_obj.day}"


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


# ===================== 抓取：icauto =====================
def fetch_city_history(city_code):
    """
    爬 icauto 的历史油价表。
    url 形如：https://www.icauto.com.cn/oil/price_440300_0.html
    返回：[日期, 90号, 92号, 93号, 95号, 97号, 0号]
    """
    url = f"https://www.icauto.com.cn/oil/price_{city_code}_0.html"
    resp = requests.get(url, headers=HEADERS, timeout=20)
    resp.encoding = "utf-8"
    html = resp.text

    # 找 historyOil 表格
    m = re.search(r"historyOil['\"]>(.*?)</table>", html, re.S)
    if not m:
        return []

    table_html = m.group(1)
    tr_list = re.findall(r"<tr>(.*?)</tr>", table_html, re.S)
    res = []
    for tr in tr_list:
        td = re.findall(r"<td>(.*?)</td>", tr, re.S)
        if len(td) < 7:
            continue
        date_str = td[0].strip()
        if not re.match(r"\d{4}-\d{2}-\d{2}", date_str):
            continue
        p90 = td[1].strip()
        p92 = td[2].strip()
        p93 = td[3].strip()
        p95 = td[4].strip()
        p97 = td[5].strip()
        p0  = td[6].strip()
        res.append([date_str, p90, p92, p93, p95, p97, p0])
    return res


def to_float(v):
    """字符串转浮点数；空、0 返回 None"""
    try:
        s = str(v).strip()
        if s == "":
            return None
        f = float(s)
        return None if f == 0 else f
    except (ValueError, TypeError):
        return None


def build_records(raw_rows):
    """
    icauto 原始行 → 我们内部结构：
    [日期, 90, 92, 93, 95, 97, 0]
    只保留 92、95、0（89 和 98 留空）
    """
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


# ===================== 合并 =====================
def merge_into_excel(sheet_name, new_records):
    """
    ★ 核心原则：
    1. 读旧数据，重命名列
    2. 旧数据每行解析出 date 对象
    3. 遍历爬虫新数据：
       - 若新日期的 ±2 天内，旧数据已有相同油价（全部共同油品一致）→ 视为已存在，跳过
       - 否则，作为新记录追加
    4. 只把"未匹配到旧记录"的新数据追加到旧数据后面
    5. 旧数据原封不动，只追加新日期
    """
    try:
        # 1. 读旧 sheet
        df_old = pd.read_excel(HISTORY_FILE, sheet_name=sheet_name)
        df_old.columns = df_old.columns.str.strip()
        print(f"DEBUG[{sheet_name}]: 读取旧sheet，原始行数={len(df_old)}, 列名={list(df_old.columns)}")

        # 2. 列名重命名
        df_old = df_old.rename(columns=MAP_EXCEL_TO_INNER)
        print(f"DEBUG[{sheet_name}]: 重命名后列名={list(df_old.columns)}")

        if "日期" not in df_old.columns:
            print(f"WARNING[{sheet_name}]: 找不到【日期】列，放弃本次合并，保护原数据！")
            return len(df_old), 0

        # 3. 旧数据逐行解析
        old_list = []
        for _, row in df_old.iterrows():
            d_obj = parse_date_to_obj(row["日期"])
            old_list.append({"date_obj": d_obj, "row": row})

        # 4. 遍历新数据，找出"需要追加"的行
        new_df = pd.DataFrame(new_records)
        new_keep = []

        for _, new_row in new_df.iterrows():
            new_d_obj = parse_date_to_obj(new_row["日期"])
            if new_d_obj is None:
                continue

            match_found = False
            for old_item in old_list:
                old_d_obj = old_item["date_obj"]
                if old_d_obj is None:
                    continue

                day_diff = abs((new_d_obj - old_d_obj).days)
                if day_diff > DAY_TOLERANCE:
                    continue

                # 找到日期接近的旧行，比油价
                old_row = old_item["row"]
                compare_cols = []
                price_all_match = True
                for col in OIL_COLS:
                    nv = new_row[col]
                    ov = old_row.get(col)
                    if pd.notna(nv) and pd.notna(ov):
                        compare_cols.append(col)
                        if abs(nv - ov) >= 0.001:
                            price_all_match = False
                if len(compare_cols) == 0:
                    continue
                if price_all_match:
                    match_found = True
                    break

            if not match_found:
                new_keep.append(new_row)

        new_only = pd.DataFrame(new_keep)
        print(f"DEBUG[{sheet_name}]: 抓 {len(new_df)} 条，过滤后需新增 {len(new_only)} 条")

        # 5. 没有新增就不动 Excel
        if new_only.empty:
            print(f"DEBUG[{sheet_name}]: 无新增数据，Excel 不动")
            return len(df_old), 0

        # 6. 合并（旧数据在前，新数据追加）
        df_all = pd.concat([df_old, new_only], ignore_index=True)

        # 7. 按日期降序
        df_all["_sort"] = df_all["日期"].apply(parse_date_to_obj)
        df_all = df_all.sort_values("_sort", ascending=False).reset_index(drop=True)
        df_all.drop(columns=["_sort"], errors="ignore", inplace=True)

        # 8. 把日期列格式统一为 YYYY/M/D
        df_all["日期"] = df_all["日期"].apply(
            lambda d: date_obj_to_str(parse_date_to_obj(d)) or d
        )

        # 9. 还原 Excel 原始列名
        INNER_TO_EXCEL = {
            "日期":     "日期",
            "89号汽油": "89汽油",
            "92号汽油": "92汽油",
            "95号汽油": "95汽油",
            "98号汽油": "98汽油",
            "0号柴油":  "0柴油",
        }
        df_all = df_all.rename(columns=INNER_TO_EXCEL)

        # 10. 写回
        with pd.ExcelWriter(HISTORY_FILE, engine="openpyxl",
                            mode="a", if_sheet_exists="replace") as writer:
            df_all.to_excel(writer, sheet_name=sheet_name, index=False)

        return len(df_all), len(new_only)

    except Exception as e:
        print(f"WARNING[{sheet_name}]: 合并异常 {e}，放弃本次合并，保护原数据！")
        return 0, 0


# ===================== 主流程 =====================
def main():
    if already_fetched_today():
        print("→ 今日已经抓取过，直接退出")
        return

    print("→ 开始爬取 17 个独立定价市")
    for sheet_name, city_name, city_code in CITIES:
        try:
            raw = fetch_city_history(city_code)
            if not raw:
                print(f"  ⚠ {sheet_name}: 未获取到数据")
                continue
            recs = build_records(raw)
            total, added = merge_into_excel(sheet_name, recs)
            print(f"  ✅ {sheet_name}: 抓 {len(recs)} 条，新增 {added} 条，现有 {total} 条")
        except Exception as e:
            print(f"  ❌ {sheet_name}: {e}")
        time.sleep(1)

    mark_fetched_today()
    print("→ 全部完成，标记今日已抓取")


if __name__ == "__main__":
    main()