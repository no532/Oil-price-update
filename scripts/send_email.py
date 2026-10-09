#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日油价推送 + 越南柴油阈值预警
跌破 25000 = 绿色（好消息）
回升超 25000 = 红色（坏消息）
通过 163 邮箱 SMTP 发送
"""

import os
import json
import smtplib
import requests
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from datetime import datetime, timedelta, timezone

# ==================== 配置 ====================
MAIL_USERNAME = os.environ.get("MAIL_USERNAME")
MAIL_PASSWORD = os.environ.get("MAIL_PASSWORD")
SMTP_HOST = "smtp.163.com"
SMTP_PORT = 465

RECEIVER_EMAIL = "Lissie.Liu@luxshare-ict.com"

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

# ==================== 预警状态 ====================
def load_alert_state():
    if os.path.exists(ALERT_STATE_FILE):
        try:
            with open(ALERT_STATE_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[WARN] 读取状态文件失败: {e}")
    return {
        "vn_diesel_below_threshold": False,
        "last_alert_date": None,
        "threshold": VN_DIESEL_THRESHOLD,
    }

def save_alert_state(state):
    try:
        with open(ALERT_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
        print(f"[INFO] 状态已保存: {state}")
    except Exception as e:
        print(f"[WARN] 保存状态失败: {e}")

def check_vn_diesel_alert(cur_price, state, today):
    """
    返回 (alert_type, message)
    alert_type: None / "trigger" / "recover"
      "trigger" = 跌破阈值 → 绿色 ✅
      "recover" = 回升超过阈值 → 红色 ⚠️
    """
    below = cur_price < VN_DIESEL_THRESHOLD
    was_below = state.get("vn_diesel_below_threshold", False)
    last_date = state.get("last_alert_date")

    # 每天只发一次
    if last_date == today:
        return None, None

    # 跌破阈值（绿色，好消息）
    if below and not was_below:
        return "trigger", (
            f"越南柴油 DO 0.05S（区域 1）当前价 <b>{fmt_int(cur_price)} VND/L</b>"
            f"，已跌破阈值 <b>{fmt_int(VN_DIESEL_THRESHOLD)} VND/L</b>"
        )

    # 从下方回升超过阈值（红色，坏消息）
    if not below and was_below:
        return "recover", (
            f"越南柴油 DO 0.05S（区域 1）当前价 <b>{fmt_int(cur_price)} VND/L</b>"
            f"，已回升超过阈值 <b>{fmt_int(VN_DIESEL_THRESHOLD)} VND/L</b>"
        )

    return None, None

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
        rows.append({
            "fuel": fuel_name,
            "retail": avg_retail,
        })
    return {
        "date": d.get("date", "—"),
        "update_time": d.get("update_time", "—"),
        "rows": rows,
    }

# ==================== 3. 国际 ====================
def fetch_international():
    d = load_json(f"{BASE_URL}/国际/oil_minline.json",
                  os.path.join(LOCAL_BASE, "国际", "oil_minline.json"))
    if not d:
        return None

    rows = []
    for name, pts in (d.get("series") or {}).items():
        if not pts:
            continue
        cur = pts[-1][1]
        meta = (d.get("meta") or {}).get(name, {})
        open_ = meta.get("open")
        prev  = meta.get("prev_close")
        diff = None
        pct = None
        if open_ and open_ != 0:
            diff = cur - open_
            pct = diff / open_ * 100
        rows.append({
            "name": name,
            "cur": cur,
            "open": open_,
            "prev": prev,
            "diff": diff,
            "pct": pct,
        })
    return {"update_time": d.get("update_time", "—"), "rows": rows}

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
        if not p:
            continue
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

# ==================== 汇总 ====================
def build_summary(domestic, hk, intl, vn):
    items = []

    if domestic and domestic.get("avg_p0") is not None:
        d0_vals = [r["d0"] for r in domestic["rows"] if r["d0"] is not None]
        avg_d0 = sum(d0_vals) / len(d0_vals) if d0_vals else None
        pct = (avg_d0 / (domestic["avg_p0"] - avg_d0) * 100) if avg_d0 and domestic["avg_p0"] else None
        items.append({
            "label": "国内柴油（全国 31 省平均）",
            "unit": "元/升",
            "value": domestic["avg_p0"],
            "diff": avg_d0,
            "pct": pct,
        })

    if hk:
        for r in hk["rows"]:
            if "柴油" in r["fuel"]:
                items.append({
                    "label": "香港柴油（平均零售价）",
                    "unit": "HKD/升",
                    "value": r["retail"],
                    "diff": None,
                    "pct": None,
                })
                break

    if intl:
        for r in intl["rows"]:
            items.append({
                "label": r["name"],
                "unit": "美元/桶",
                "value": r["cur"],
                "diff": r["diff"],
                "pct": r["pct"],
            })

    if vn:
        for r in vn["rows"]:
            if r["prefix"] == "DO005S_M2":
                items.append({
                    "label": "越南柴油 0.05S（区域 1）",
                    "unit": "VND/升",
                    "value": r["z1"],
                    "diff": None,
                    "pct": None,
                })
                break

    html = """
    <div style="background:#fff8f7;border:1px solid #f5d0cc;border-radius:8px;
                padding:16px 20px;margin-bottom:20px;">
      <h3 style="color:#c9302c;margin:0 0 12px;font-size:16px;">📊 今日汇总</h3>
      <table cellpadding="6" cellspacing="0" style="font-size:13px;width:100%;">
    """
    for it in items:
        diff_str = "—"
        if it["diff"] is not None:
            color = color_of(it["diff"])
            pct_str = fmt_signed_pct(it["pct"]) if it["pct"] is not None else "—"
            diff_str = (
                f"<span style='color:{color};'>"
                f"{fmt_signed(it['diff'])}（{pct_str}）"
                f"</span>"
            )
        val = it["value"]
        val_str = fmt_num(val) if it["unit"] != "VND/升" else fmt_int(val)
        html += (
            f"<tr>"
            f"<td style='color:#5a6b7b;'>{it['label']}</td>"
            f"<td align='right' style='font-weight:600;'>"
            f"{val_str} <span style='color:#8492a6;font-weight:400;'>{it['unit']}</span></td>"
            f"<td align='right' style='width:180px;'>较上次：{diff_str}</td>"
            f"</tr>"
        )
    html += "</table></div>"
    return html

# ==================== 预警横幅（颜色已交换） ====================
def build_alert_banner(alert_type, message):
    if not alert_type:
        return ""
    # 跌破 = 好消息 = 绿色
    if alert_type == "trigger":
        bg, border, color, icon = "#eef9f1", "#b7e4c7", "#27ae60", "✅"
    # 回升 = 坏消息 = 红色
    else:  # "recover"
        bg, border, color, icon = "#fdf0ef", "#f5c6c0", "#c9302c", "⚠️"
    return f"""
    <div style="background:{bg};border:2px solid {border};border-radius:8px;
                padding:16px 20px;margin-bottom:20px;">
      <div style="font-size:16px;font-weight:700;color:{color};">
        {icon} 油价预警
      </div>
      <div style="font-size:14px;color:#1f2d3d;margin-top:8px;line-height:1.7;">
        {message}
      </div>
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
    if not domestic:
        return "<p style='color:#e74c3c'>国内数据读取失败</p>"
    rows = domestic["rows"]
    html = f"""
    <p style="color:#8492a6;font-size:12px;">共 {len(rows)} 个省市</p>
    <table border="1" cellpadding="8" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;width:100%;">
      <thead style="background:#fafbfc;">
        <tr>
          <th align="left">省份</th><th>日期</th>
          <th>92# (元/升)</th><th>95# (元/升)</th><th>0# 柴油 (元/升)</th>
        </tr>
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
    if not d:
        return "<p style='color:#e74c3c'>香港数据读取失败</p>"
    html = f"""
    <p style="color:#8492a6;font-size:12px;">
      数据日期：{d['date']}　更新时间：{d['update_time']}
    </p>
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
    if not d:
        return "<p style='color:#e74c3c'>国际原油数据读取失败</p>"
    html = f"""
    <p style="color:#8492a6;font-size:12px;">更新时间：{d['update_time']}</p>
    <table border="1" cellpadding="8" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;width:100%;">
      <thead style="background:#fafbfc;">
        <tr>
          <th align="left">品种</th><th>当前 (美元/桶)</th>
          <th>较今开</th><th>今开</th><th>昨收</th>
        </tr>
      </thead><tbody>
    """
    for r in d["rows"]:
        if r["diff"] is not None:
            color = color_of(r["diff"])
            pct_str = (
                f"<span style='color:{color}'>"
                f"{fmt_signed(r['diff'])}（{fmt_signed_pct(r['pct'])}）</span>"
            )
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
    if not d:
        return "<p style='color:#e74c3c'>越南数据读取失败</p>"
    html = f"""
    <p style="color:#8492a6;font-size:12px;">
      数据日期：{d['date']}　更新时间：{d['update_time']}
    </p>
    <table border="1" cellpadding="8" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;width:100%;">
      <thead style="background:#fafbfc;">
        <tr>
          <th align="left">油品</th>
          <th>区域 1 (VND/L)</th><th>区域 2 (VND/L)</th>
        </tr>
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
def build_email(alert_type=None, alert_msg=None):
    today = bj_now().strftime("%Y-%m-%d")

    domestic = fetch_domestic()
    hk = fetch_hongkong()
    intl = fetch_international()
    vn = fetch_vietnam()

    summary_html = build_summary(domestic, hk, intl, vn)
    alert_html = build_alert_banner(alert_type, alert_msg)

    # 邮件标题（颜色对应的图标）
    if alert_type == "trigger":
        subject = f"✅ 越南柴油跌破阈值 · {today}"
    elif alert_type == "recover":
        subject = f"⚠️ 越南柴油回升 · {today}"
    else:
        subject = f"【每日油价】{today}"

    body = f"""
    <html><body style="font-family:Arial,'Microsoft YaHei',sans-serif;
                        color:#1f2d3d;max-width:960px;margin:auto;padding:16px;">
      <h2 style="color:#c9302c;text-align:center;">每日油价汇总 · {today}</h2>
      <p style="color:#8492a6;font-size:12px;text-align:center;">
        自动推送，数据仅供参考
      </p>

      {alert_html}
      {summary_html}

      {section("一、国内各省市油价", render_domestic(domestic))}
      {section("二、香港油价",       render_hongkong(hk))}
      {section("三、国际原油",       render_international(intl))}
      {section("四、越南油价",       render_vietnam(vn))}

      <hr style="border:none;border-top:1px dashed #e4e7ed;margin:24px 0;">
      <div style="text-align:center;margin:20px 0;">
        <a href="{BASE_URL}" style="display:inline-block;padding:10px 24px;
           background:#c9302c;color:#fff;text-decoration:none;border-radius:6px;
           font-size:14px;font-weight:600;">
          🌐 查看完整网页版
        </a>
      </div>
      <p style="color:#8492a6;font-size:12px;text-align:center;">
        或复制链接：<a href="{BASE_URL}" style="color:#409eff;">{BASE_URL}</a><br>
        本邮件由 GitHub Actions 自动生成，请勿回复
      </p>
    </body></html>
    """
    return today, subject, body

# ==================== 发送 ====================
def send_mail(subject, html):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = MAIL_USERNAME
    msg["To"] = RECEIVER_EMAIL
    msg.attach(MIMEText(html, "html", "utf-8"))

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(MAIL_USERNAME, MAIL_PASSWORD)
        server.sendmail(MAIL_USERNAME, [RECEIVER_EMAIL], msg.as_string())

def main():
    if not MAIL_USERNAME or not MAIL_PASSWORD:
        raise SystemExit("缺少 MAIL_USERNAME 或 MAIL_PASSWORD")

    today = bj_now().strftime("%Y-%m-%d")

    vn = fetch_vietnam()
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
        print(f"[INFO] 越南柴油 05S 区域1 = {vn_diesel_z1}, 预警 = {alert_type}")

    _, subject, html = build_email(alert_type, alert_msg)
    try:
        send_mail(subject, html)
        print("邮件发送成功 →", RECEIVER_EMAIL)
    except Exception as e:
        print("邮件发送失败:", e)
        raise

    if alert_type:
        state["vn_diesel_below_threshold"] = (vn_diesel_z1 < VN_DIESEL_THRESHOLD)
        state["last_alert_date"] = today
    else:
        state["vn_diesel_below_threshold"] = (vn_diesel_z1 is not None
                                              and vn_diesel_z1 < VN_DIESEL_THRESHOLD)
    state["threshold"] = VN_DIESEL_THRESHOLD
    save_alert_state(state)

if __name__ == "__main__":
    main()
