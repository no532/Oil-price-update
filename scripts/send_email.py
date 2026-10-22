#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日油价推送 + 越南柴油预警
汇总页含：当前价、上次变动日期、下次调整日期（仅国内）
通过 163 邮箱 SMTP 发送（一封邮件，多个收件人）
"""

import os
import csv
import json
import smtplib
import requests
import openpyxl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from datetime import datetime, timedelta, timezone

# ==================== 配置 ====================
MAIL_USERNAME = os.environ.get("MAIL_USERNAME")
MAIL_PASSWORD = os.environ.get("MAIL_PASSWORD")
SMTP_HOST = "smtp.163.com"
SMTP_PORT = 465

def parse_emails(env_name):
    raw = os.environ.get(env_name, "")
    return [e.strip() for e in raw.split(",") if e.strip()]

DAILY_RECEIVERS = parse_emails("DAILY_RECEIVERS")
ALERT_RECEIVERS = parse_emails("ALERT_RECEIVERS")

BASE_URL = "https://no532.github.io/Oil-price-update"
LOCAL_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

VN_DIESEL_THRESHOLD = 25000
ALERT_STATE_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "alert_state.json"
)

# ==================== 时间 ====================
def bj_now():
    return datetime.now(timezone(timedelta(hours=8)))

# ==================== 工具 ====================
def http_get_json(url, timeout=15):
    try:
        r = requests.get(url, timeout=timeout, headers={
            "User-Agent": "Mozilla/5.0 (compatible; FuelBot/1.0)"
        })
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"[WARN] 拉取失败 {url}: {e}")
        return None

def load_json_local(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"[WARN] 本地读取失败 {path}: {e}")
        return None

def load_json(url, local):
    d = http_get_json(url)
    if d is not None:
        return d
    return load_json_local(local)
    
def round2(v):
    if v is None:
        return None
    return round(v * 100) / 100
    
def fmt_num(v):
    if v is None: return "—"
    try: return f"{float(v):.2f}"
    except: return "—"

def fmt_int(v):
    if v is None: return "—"
    try: return f"{int(v):,}"
    except: return "—"

def fmt_signed(v):
    if v is None: return "—"
    try:
        f = float(v)
        return f"{f:+.2f}" if f >= 0 else f"{f:.2f}"
    except: return "—"

def fmt_signed_pct(v):
    if v is None: return "—"
    try:
        f = float(v)
        return f"{f:+.2f}%" if f >= 0 else f"{f:.2f}%"
    except: return "—"

def color_of(v):
    if v is None: return "#8492a6"
    try:
        f = float(v)
        if f > 0: return "#e74c3c"
        if f < 0: return "#27ae60"
    except: pass
    return "#8492a6"

def parse_excel_date(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, str):
        for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d"):
            try:
                return datetime.strptime(v.strip(), fmt)
            except: pass
    return None

# ==================== 读调价日历 ====================
def load_price_calendar():
    path = os.path.join(LOCAL_BASE, "国内", "油价调整日历.xlsx")
    if not os.path.exists(path):
        print(f"[WARN] 找不到 {path}")
        return []

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["Sheet1"]
        dates = []
        for row in range(2, ws.max_row + 1):
            v = ws.cell(row=row, column=1).value
            if v is None: continue
            d = parse_excel_date(v)
            if d:
                # ★ 只要第一列有日期，统统算作调价日
                dates.append(d)
        dates.sort()
        if dates:
            print(f"[INFO] 调价日历读取 {len(dates)} 条：{dates[0]} ~ {dates[-1]}")
        return dates
    except Exception as e:
        print(f"[WARN] 读调价日历失败: {e}")
        return []

def next_adjust_date(calendar, today):
    for d in calendar:
        if d.date() >= today.date():
            return d
    return None

# ==================== 读各省历史油价 ====================
def load_province_history():
    path = os.path.join(LOCAL_BASE, "国内", "各省历史油价数据.xlsx")
    if not os.path.exists(path):
        print(f"[WARN] 找不到 {path}")
        return {}

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        result = {}
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = []
            for row in range(2, ws.max_row + 1):
                d = parse_excel_date(ws.cell(row=row, column=1).value)
                v = ws.cell(row=row, column=6).value  # F 列 = 0号柴油
                if d and v is not None:
                    try:
                        rows.append((d, float(v)))
                    except: pass
            if rows:
                rows.sort(key=lambda x: x[0], reverse=True)
                result[sheet_name] = rows
        print(f"[INFO] 各省历史油价读取 {len(result)} 个省")
        return result
    except Exception as e:
        print(f"[WARN] 读各省历史油价失败: {e}")
        return {}

def find_last_change_date(history_rows, current_price):
    if not history_rows or current_price is None:
        return None, None
    for d, v in history_rows:
        if abs(v - current_price) > 0.01:
            return d, v
    return None, None

# ==================== 读香港历史 CSV ====================
def load_hk_history():
    path = os.path.join(LOCAL_BASE, "香港", "香港历史油价.csv")
    if not os.path.exists(path):
        print(f"[WARN] 找不到 {path}")
        return []

    try:
        with open(path, encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            rows = []
            for r in reader:
                d = r.get("Date", "").strip()
                try:
                    dt = datetime.strptime(d, "%Y/%m/%d")
                except: continue
                vals = []
                for col in ["中石化 (零售牌價)", "中國石油 (零售牌價)",
                            "加德士 (零售牌價)", "埃索 (零售牌價)", "蜆殼 (零售牌價)"]:
                    try:
                        vals.append(float(r[col]))
                    except: pass
                if vals:
                    rows.append((dt, sum(vals) / len(vals)))
            rows.sort(key=lambda x: x[0], reverse=True)
            print(f"[INFO] 香港历史读取 {len(rows)} 条")
            return rows
    except Exception as e:
        print(f"[WARN] 读香港历史失败: {e}")
        return []

# ==================== 读越南历史 CSV ====================
def load_vn_history():
    path = os.path.join(LOCAL_BASE, "越南", "越南历史油价.csv")
    if not os.path.exists(path):
        print(f"[WARN] 找不到 {path}")
        return []

    try:
        with open(path, encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            rows = []
            for r in reader:
                d = r.get("Date", "").strip()
                v = r.get("DO005S_M2_Z1", "").strip()
                if not d or not v: continue
                try:
                    dt = datetime.strptime(d, "%Y/%m/%d")
                    rows.append((dt, float(v)))
                except: continue
            rows.sort(key=lambda x: x[0], reverse=True)
            print(f"[INFO] 越南历史读取 {len(rows)} 条")
            return rows
    except Exception as e:
        print(f"[WARN] 读越南历史失败: {e}")
        return []

# ==================== 1. 国内 ====================
def fetch_domestic():
    d = load_json(f"{BASE_URL}/国内/data.json",
                  os.path.join(LOCAL_BASE, "国内", "data.json"))
    if not d:
        return None

    rows = []
    for province, arr in d.items():
        if not arr or not isinstance(arr, list):
            continue
        latest = arr[0]
        rows.append({
            "province": province,
            "date": latest.get("date", "—"),
            "p92": latest.get("92号汽油"),
            "p95": latest.get("95号汽油"),
            "p0":  latest.get("0号柴油"),
            "d0":  latest.get("0号柴油_diff"),
        })
    rows.sort(key=lambda x: x["province"])
    p0_vals = [r["p0"] for r in rows if r["p0"] is not None]
    avg_p0 = sum(p0_vals) / len(p0_vals) if p0_vals else None
    return {"rows": rows, "avg_p0": avg_p0}

# ==================== 2. 香港 ====================
def fetch_hongkong():
    d = load_json(f"{BASE_URL}/香港/hk_fuel.json",
                  os.path.join(LOCAL_BASE, "香港", "hk_fuel.json"))
    if not d:
        return None

    open_data = d.get("open_data", {})
    rows = []
    for fuel_name, brands in open_data.items():
        if not isinstance(brands, dict):
            continue
        vals = [v for v in brands.values() if isinstance(v, (int, float))]
        avg_retail = sum(vals) / len(vals) if vals else None
        rows.append({"fuel": fuel_name, "retail": avg_retail})
    return {
        "date": d.get("date", "—"),
        "update_time": d.get("update_time", "—"),
        "rows": rows,
    }

# ==================== 3. 国际 ====================
def fetch_international():
    d = load_json(f"{BASE_URL}/国际/oil_minline.json",
                  os.path.join(LOCAL_BASE, "国际", "oil_minline.json"))
    if not d: return None

    rows = []
    for name, pts in (d.get("series") or {}).items():
        if not pts: continue
        
        # ★ 1. 取今开 (Open)
        meta = (d.get("meta") or {}).get(name, {})
        open_price = meta.get("open")
        
        # 如果今开为空，退而求其次取最新分时价
        cur = open_price if open_price else pts[-1][1]
        
        rows.append({
            "name": name, 
            "cur": cur,     # 这里的 cur 其实代表今天的开盘价
            "open": open_price, 
            "diff": None,   # 波动留给后面用日K算
            "pct": None
        })
    return {"update_time": d.get("update_time", "—"), "rows": rows}

def load_intl_monthly():
    d = load_json(f"{BASE_URL}/国际/oil_monthly.json",
                  os.path.join(LOCAL_BASE, "国际", "oil_monthly.json"))
    if not d:
        return {}
    result = {}
    for name, arr in (d.get("series") or {}).items():
        rows = []
        for item in arr:
            try:
                dt = datetime.strptime(item[0], "%Y-%m-%d")
                rows.append((dt, float(item[1])))
            except: pass
        rows.sort(key=lambda x: x[0], reverse=True)
        result[name] = rows
    return result

# ==================== 4. 越南 ====================
def fetch_vietnam():
    d = load_json(f"{BASE_URL}/越南/vietnam_fuel.json",
                  os.path.join(LOCAL_BASE, "越南", "vietnam_fuel.json"))
    if not d:
        return None

    order = ["DO005S_M2","DO001S_M5","E10RON95_M5","E10RON95_M3","E5RON92_M2","Kero2K"]
    names = d.get("fuel_names", {})
    prices = d.get("prices", {})

    rows = []
    for k in order:
        p = prices.get(k)
        if not p: continue
        rows.append({
            "prefix": k,
            "name": names.get(k, k),
            "z1": p.get("Z1"),
            "z2": p.get("Z2"),
        })
    return {
        "date": d.get("date", "—"),
        "update_time": d.get("update_time", "—"),
        "rows": rows,
    }

# ==================== 预警状态 ====================
def load_alert_state():
    if os.path.exists(ALERT_STATE_FILE):
        try:
            with open(ALERT_STATE_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[WARN] 读取状态文件失败: {e}")
    return {"vn_diesel_below_threshold": False,
            "last_alert_date": None,
            "threshold": VN_DIESEL_THRESHOLD}

def save_alert_state(state):
    try:
        with open(ALERT_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
        print(f"[INFO] 状态已保存: {state}")
    except Exception as e:
        print(f"[WARN] 保存状态失败: {e}")

def check_vn_diesel_alert(cur_price, state, today):
    below = cur_price < VN_DIESEL_THRESHOLD
    was_below = state.get("vn_diesel_below_threshold", False)
    if state.get("last_alert_date") == today:
        return None, None

    if below and not was_below:
        return "trigger", (
            f"越南柴油 DO 0.05S（区域 1）当前价 <b>{fmt_int(cur_price)} VND/L</b>"
            f"，已跌破阈值 <b>{fmt_int(VN_DIESEL_THRESHOLD)} VND/L</b>"
        )
    if not below and was_below:
        return "recover", (
            f"越南柴油 DO 0.05S（区域 1）当前价 <b>{fmt_int(cur_price)} VND/L</b>"
            f"，已回升超过阈值 <b>{fmt_int(VN_DIESEL_THRESHOLD)} VND/L</b>"
        )
    return None, None

# ==================== 汇总表 ====================
def build_summary(domestic, hk, intl, vn,
                  prov_history, hk_history, vn_history, intl_history,
                  calendar):
    today = bj_now()
    rows = []

    # --- 国内柴油 ---
    # --- 国内柴油 ---
    if domestic and domestic["avg_p0"] is not None:
        avg = domestic["avg_p0"]
        
        d0_vals = [r["d0"] for r in domestic["rows"] if r["d0"] is not None]
        avg_d0 = sum(d0_vals) / len(d0_vals) if d0_vals else 0
        
        diff = round2(avg_d0)
        pct = round2(diff / (avg - diff) * 100) if avg and diff else 0.0
        
        # ★ 核心修改：找“上一个”调价日
        last_date_str = "—"
        past_dates = [d for d in calendar if d.date() <= today.date()]
        # 如果今天是在调价日（10/16）之后，past_dates 最后一个就是 10/16，倒数第二个就是 9/25
        if len(past_dates) >= 2:
            last_date_str = past_dates[-2].strftime('%Y-%m-%d')
        elif len(past_dates) == 1:
            last_date_str = past_dates[-1].strftime('%Y-%m-%d')
            
        # 下次调整日期（保持不变）
        next_d = next_adjust_date(calendar, today)
        if next_d:
            days = (next_d.date() - today.date()).days
            if days == 0:
                next_str = f"<b style='color:#e74c3c;'>今天调整</b>"
            else:
                next_str = f"{next_d.strftime('%Y-%m-%d')}（还有 {days} 天）"
        else:
            next_str = "—"

        rows.append({
            "label": "国内柴油（31 省平均）",
            "value": f"{fmt_num(avg)}<span class='u'>元/升</span>",
            "last_date": last_date_str,
            "diff": diff,
            "pct": pct,
            "next": next_str,
        })

    # --- 香港柴油 ---
    if hk:
        for r in hk["rows"]:
            if "柴油" in r["fuel"]:
                cur = r["retail"]
                last_date, last_price = find_last_change_date(hk_history, cur)
                diff = None; pct = None; last_date_str = "—"
                if last_date and last_price:
                    diff = cur - last_price
                    pct = diff / last_price * 100 if last_price else None
                    last_date_str = last_date.strftime('%Y-%m-%d')
                rows.append({
                    "label": "香港柴油（平均零售价）",
                    "value": f"{fmt_num(cur)}<span class='u'>HKD/升</span>",
                    "last_date": last_date_str,
                    "diff": diff,
                    "pct": pct,
                    "next": "—",
                })
                break

    # --- 国际 WTI / 布伦特 ---
    if intl:
        for r in intl["rows"]:
            cur = r["cur"]
            hist = intl_history.get(r["name"], [])
            last_date, last_price = find_last_change_date(hist, cur)
            diff = None; pct = None; last_date_str = "—"
            if last_date and last_price:
                diff = cur - last_price
                pct = diff / last_price * 100 if last_price else None
                last_date_str = last_date.strftime('%Y-%m-%d')
            rows.append({
                "label": r["name"],
                "value": f"{fmt_num(cur)}<span class='u'>美元/桶</span>",
                "last_date": last_date_str,
                "diff": diff,
                "pct": pct,
                "next": "—",
            })

    # --- 越南柴油 05S ---
    if vn:
        for r in vn["rows"]:
            if r["prefix"] == "DO005S_M2":
                cur = r["z1"]
                last_date, last_price = find_last_change_date(vn_history, cur)
                diff = None; pct = None; last_date_str = "—"
                if last_date and last_price:
                    diff = cur - last_price
                    pct = diff / last_price * 100 if last_price else None
                    last_date_str = last_date.strftime('%Y-%m-%d')
                rows.append({
                    "label": "越南柴油 0.05S（区域 1）",
                    "value": f"{fmt_int(cur)}<span class='u'>VND/升</span>",
                    "last_date": last_date_str,
                    "diff": diff,
                    "pct": pct,
                    "next": "—",
                })
                break

    # 拼 HTML 表格
    html = """
    <div style="background:#fff8f7;border:1px solid #f5d0cc;border-radius:8px;
                padding:16px 20px;margin-bottom:20px;">
      <h3 style="color:#c9302c;margin:0 0 12px;font-size:16px;">📊 今日汇总</h3>
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <thead>
          <tr style="background:#fbeae8;">
            <th style="padding:8px 10px;text-align:left;color:#c9302c;font-weight:600;border-bottom:1px solid #f0d0cc;">品种</th>
            <th style="padding:8px 10px;text-align:right;color:#c9302c;font-weight:600;border-bottom:1px solid #f0d0cc;">当前价</th>
            <th style="padding:8px 10px;text-align:center;color:#c9302c;font-weight:600;border-bottom:1px solid #f0d0cc;">较上次</th>
            <th style="padding:8px 10px;text-align:center;color:#c9302c;font-weight:600;border-bottom:1px solid #f0d0cc;">下次调整</th>
          </tr>
        </thead>
        <tbody>
    """

    for r in rows:
        if r["diff"] is not None and r["pct"] is not None:
            c = color_of(r["diff"])
            sign = "+" if r["diff"] > 0 else ""
            diff_str = (
                f"<div style='color:{c};font-weight:600;'>"
                f"{sign}{fmt_num(r['diff'])}（{sign}{r['pct']:.2f}%）"
                f"</div>"
                f"<div style='color:#8492a6;font-size:11px;margin-top:2px;'>"
                f"较 {r['last_date']}"
                f"</div>"
            )
        else:
            diff_str = (
                f"<div style='color:#8492a6;'>—</div>"
                f"<div style='color:#8492a6;font-size:11px;margin-top:2px;'>较 —</div>"
            )

        html += f"""
        <tr>
          <td style="padding:10px;text-align:left;border-bottom:1px dashed #f0d0cc;color:#5a6b7b;">{r['label']}</td>
          <td style="padding:10px;text-align:right;border-bottom:1px dashed #f0d0cc;font-weight:600;color:#1f2d3d;">{r['value']}</td>
          <td style="padding:10px;text-align:center;border-bottom:1px dashed #f0d0cc;">{diff_str}</td>
          <td style="padding:10px;text-align:center;border-bottom:1px dashed #f0d0cc;color:#5a6b7b;font-size:12px;">{r['next']}</td>
        </tr>
        """

    html += """
        </tbody>
      </table>
      <style>
        .u { font-size:11px; color:#8492a6; font-weight:400; margin-left:2px; }
      </style>
    </div>
    """
    return html

# ==================== 预警横幅 ====================
def build_alert_banner(alert_type, message):
    if not alert_type: return ""
    if alert_type == "trigger":
        bg, border, color, icon = "#eef9f1", "#b7e4c7", "#27ae60", "✅"
    else:
        bg, border, color, icon = "#fdf0ef", "#f5c6c0", "#c9302c", "⚠️"
    return f"""
    <div style="background:{bg};border:2px solid {border};border-radius:8px;
                padding:16px 20px;margin-bottom:20px;">
      <div style="font-size:16px;font-weight:700;color:{color};">{icon} 油价预警</div>
      <div style="font-size:14px;color:#1f2d3d;margin-top:8px;line-height:1.7;">{message}</div>
    </div>
    """

# ==================== 明细 ====================
def section(title, content):
    return f"""
    <div style="margin:24px 0;">
      <h3 style="color:#c9302c;border-left:4px solid #c9302c;
                 padding-left:10px;margin:0 0 10px;font-size:15px;">{title}</h3>
      {content}
    </div>
    """

def render_domestic(domestic):
    if not domestic: return "<p style='color:#e74c3c'>国内数据读取失败</p>"
    rows = domestic["rows"]
    html = f"""
    <p style="color:#8492a6;font-size:12px;">共 {len(rows)} 个省市</p>
    <table border="1" cellpadding="8" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;width:100%;">
      <thead style="background:#fafbfc;">
        <tr><th align="left">省份</th><th>日期</th>
            <th>92# (元/升)</th><th>95# (元/升)</th><th>0# 柴油 (元/升)</th></tr>
      </thead><tbody>
    """
    for r in rows:
        html += (
            f"<tr><td>{r['province']}</td><td align='center'>{r['date']}</td>"
            f"<td align='right'>{fmt_num(r['p92'])}</td>"
            f"<td align='right'>{fmt_num(r['p95'])}</td>"
            f"<td align='right'>{fmt_num(r['p0'])}</td></tr>"
        )
    html += "</tbody></table>"
    return html

def render_hongkong(d):
    if not d: return "<p style='color:#e74c3c'>香港数据读取失败</p>"
    html = f"""
    <p style="color:#8492a6;font-size:12px;">数据日期：{d['date']}　更新时间：{d['update_time']}</p>
    <table border="1" cellpadding="8" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;width:100%;">
      <thead style="background:#fafbfc;">
        <tr><th align="left">油品</th><th>平均零售价 (HKD/L)</th></tr>
      </thead><tbody>
    """
    for r in d["rows"]:
        html += f"<tr><td>{r['fuel']}</td><td align='right'>{fmt_num(r['retail'])}</td></tr>"
    html += "</tbody></table>"
    return html

def render_international(d):
    if not d: return "<p style='color:#e74c3c'>国际原油数据读取失败</p>"
    html = f"""
    <p style="color:#8492a6;font-size:12px;">更新时间：{d['update_time']}</p>
    <table border="1" cellpadding="8" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;width:100%;">
      <thead style="background:#fafbfc;">
        <tr><th align="left">品种</th><th>当前 (美元/桶)</th>
            <th>较今开</th><th>今开</th><th>昨收</th></tr>
      </thead><tbody>
    """
    for r in d["rows"]:
        if r["diff"] is not None:
            color = color_of(r["diff"])
            pct_str = f"<span style='color:{color}'>{fmt_signed(r['diff'])}（{fmt_signed_pct(r['pct'])}）</span>"
        else:
            pct_str = "—"
        html += (
            f"<tr><td>{r['name']}</td>"
            f"<td align='right'>{fmt_num(r['cur'])}</td>"
            f"<td align='center'>{pct_str}</td>"
            f"<td align='right'>{fmt_num(r['open'])}</td>"
            f"<td align='right'>{fmt_num(r['prev'])}</td></tr>"
        )
    html += "</tbody></table>"
    return html

def render_vietnam(d):
    if not d: return "<p style='color:#e74c3c'>越南数据读取失败</p>"
    html = f"""
    <p style="color:#8492a6;font-size:12px;">数据日期：{d['date']}　更新时间：{d['update_time']}</p>
    <table border="1" cellpadding="8" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;width:100%;">
      <thead style="background:#fafbfc;">
        <tr><th align="left">油品</th>
            <th>区域 1 (VND/L)</th><th>区域 2 (VND/L)</th></tr>
      </thead><tbody>
    """
    for r in d["rows"]:
        html += (
            f"<tr><td>{r['name']}</td>"
            f"<td align='right'>{fmt_int(r['z1'])}</td>"
            f"<td align='right'>{fmt_int(r['z2'])}</td></tr>"
        )
    html += "</tbody></table>"
    return html

# ==================== 拼邮件 ====================
def build_daily_email(domestic, hk, intl, vn,
                      prov_history, hk_history, vn_history, intl_history,
                      calendar):
    today = bj_now().strftime("%Y-%m-%d")
    summary_html = build_summary(domestic, hk, intl, vn,
                                 prov_history, hk_history, vn_history, intl_history,
                                 calendar)

    body = f"""
    <html><body style="font-family:Arial,'Microsoft YaHei',sans-serif;
                        color:#1f2d3d;max-width:960px;margin:auto;padding:16px;">
      <h2 style="color:#c9302c;text-align:center;">每日油价汇总 · {today}</h2>
      <p style="color:#8492a6;font-size:12px;text-align:center;">自动推送，数据仅供参考</p>

      {summary_html}

      {section("一、国内各省市油价", render_domestic(domestic))}
      {section("二、香港油价",       render_hongkong(hk))}
      {section("三、国际原油",       render_international(intl))}
      {section("四、越南油价",       render_vietnam(vn))}

      <hr style="border:none;border-top:1px dashed #e4e7ed;margin:24px 0;">
      <div style="text-align:center;margin:20px 0;">
        <a href="{BASE_URL}" style="display:inline-block;padding:10px 24px;
           background:#c9302c;color:#fff;text-decoration:none;border-radius:6px;
           font-size:14px;font-weight:600;">🌐 查看完整网页版</a>
      </div>
      <p style="color:#8492a6;font-size:12px;text-align:center;">
        本邮件由 GitHub Actions 自动生成，请勿回复
      </p>
    </body></html>
    """
    return f"【每日油价】{today}", body

def build_alert_email(alert_type, alert_msg, vn):
    today = bj_now().strftime("%Y-%m-%d")
    alert_html = build_alert_banner(alert_type, alert_msg)
    vn_html = render_vietnam(vn) if vn else ""
    body = f"""
    <html><body style="font-family:Arial,'Microsoft YaHei',sans-serif;
                        color:#1f2d3d;max-width:960px;margin:auto;padding:16px;">
      <h2 style="color:#c9302c;text-align:center;">越南柴油预警 · {today}</h2>
      <p style="color:#8492a6;font-size:12px;text-align:center;">
        仅在越南柴油 05S 跨越 25000 VND/L 阈值时推送
      </p>

      {alert_html}
      {section("越南油价明细", vn_html)}

      <hr style="border:none;border-top:1px dashed #e4e7ed;margin:24px 0;">
      <div style="text-align:center;margin:20px 0;">
        <a href="{BASE_URL}" style="display:inline-block;padding:10px 24px;
           background:#c9302c;color:#fff;text-decoration:none;border-radius:6px;
           font-size:14px;font-weight:600;">🌐 查看完整网页版</a>
      </div>
      <p style="color:#8492a6;font-size:12px;text-align:center;">
        本邮件由 GitHub Actions 自动生成，请勿回复
      </p>
    </body></html>
    """
    if alert_type == "trigger":
        subject = f"✅ 越南柴油跌破阈值 · {today}"
    else:
        subject = f"⚠️ 越南柴油回升超过阈值 · {today}"
    return subject, body

# ==================== 发送（一起发） ====================
def send_mail(subject, html, receivers):
    if not receivers:
        print("[WARN] 收件人列表为空，跳过发送")
        return
    msg = MIMEMultipart("alternative")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = MAIL_USERNAME
    msg["To"] = ", ".join(receivers)
    msg.attach(MIMEText(html, "html", "utf-8"))

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as server:
        server.login(MAIL_USERNAME, MAIL_PASSWORD)
        server.sendmail(MAIL_USERNAME, receivers, msg.as_string())

    print(f"✓ 已发送 → {receivers}")

def main():
    if not MAIL_USERNAME or not MAIL_PASSWORD:
        raise SystemExit("缺少 MAIL_USERNAME 或 MAIL_PASSWORD")

    today = bj_now().strftime("%Y-%m-%d")

    # 1. 拉所有数据
    domestic = fetch_domestic()
    hk = fetch_hongkong()
    intl = fetch_international()
    vn = fetch_vietnam()

    # 2. 拉历史数据
    prov_history = load_province_history()
    hk_history = load_hk_history()
    vn_history = load_vn_history()
    intl_history = load_intl_monthly()
    calendar = load_price_calendar()

    # 3. 判断预警
    vn_diesel_z1 = None
    if vn:
        for r in vn["rows"]:
            if r["prefix"] == "DO005S_M2":
                vn_diesel_z1 = r["z1"]
                break

    state = load_alert_state()
    alert_type = None
    alert_msg = None
    if vn_diesel_z1 is not None:
        alert_type, alert_msg = check_vn_diesel_alert(vn_diesel_z1, state, today)
        print(f"[INFO] 越南柴油 05S 区域1 = {vn_diesel_z1}, "
              f"below = {state.get('vn_diesel_below_threshold')}, "
              f"预警 = {alert_type}")

    # 4. 发每日推送
    try:
        daily_subject, daily_html = build_daily_email(
            domestic, hk, intl, vn,
            prov_history, hk_history, vn_history, intl_history,
            calendar
        )
        send_mail(daily_subject, daily_html, DAILY_RECEIVERS)
    except Exception as e:
        print("[ERROR] 每日推送失败:", e)

    # 5. 如果有预警，额外发一封
    if alert_type:
        try:
            alert_subject, alert_html = build_alert_email(alert_type, alert_msg, vn)
            send_mail(alert_subject, alert_html, ALERT_RECEIVERS)
        except Exception as e:
            print("[ERROR] 预警推送失败:", e)

    # 6. 更新预警状态
    if vn_diesel_z1 is not None:
        state["vn_diesel_below_threshold"] = (vn_diesel_z1 < VN_DIESEL_THRESHOLD)
    if alert_type:
        state["last_alert_date"] = today
    state["threshold"] = VN_DIESEL_THRESHOLD
    save_alert_state(state)

if __name__ == "__main__":
    main()
