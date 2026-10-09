#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日油价推送：
  顶部汇总（4 大板块柴油价格+涨跌）
  → 国内 / 香港 / 国际 / 越南 明细
  → 网页链接
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
MAIL_USERNAME = os.environ.get("MAIL_USERNAME")   # 如 abc@163.com
MAIL_PASSWORD = os.environ.get("MAIL_PASSWORD")   # 163 邮箱授权码
SMTP_HOST = "smtp.163.com"
SMTP_PORT = 465

RECEIVER_EMAIL = "Lissie.Liu@luxshare-ict.com"    # ★ 收件邮箱

BASE_URL = "https://no532.github.io/Oil-price-update"
LOCAL_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

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
    if v is None:
        return "—"
    try:
        return f"{float(v):.2f}"
    except Exception:
        return "—"

def fmt_int(v):
    if v is None:
        return "—"
    try:
        return f"{int(v):,}"
    except Exception:
        return "—"

def fmt_signed(v):
    if v is None:
        return "—"
    try:
        f = float(v)
        return f"{f:+.2f}" if f >= 0 else f"{f:.2f}"
    except Exception:
        return "—"

def color_of(v):
    if v is None:
        return "#8492a6"
    try:
        f = float(v)
        if f > 0: return "#e74c3c"
        if f < 0: return "#27ae60"
    except Exception:
        pass
    return "#8492a6"

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
    return rows

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
        sample_brand = "中石化" if "中石化" in brands else list(brands.keys())[0]
        retail = brands.get(sample_brand)
        rows.append({
            "fuel": fuel_name,
            "brand": sample_brand,
            "retail": retail,
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

    if domestic:
        target = None
        for r in domestic:
            if "北京" in r["province"]:
                target = r
                break
        if not target:
            target = domestic[0]
        items.append({
            "label": f"国内柴油（{target['province']}）",
            "unit": "元/升",
            "value": target.get("p0"),
            "diff": target.get("d0"),
        })

    if hk:
        for r in hk["rows"]:
            if "柴油" in r["fuel"]:
                items.append({
                    "label": f"香港柴油（{r['brand']}）",
                    "unit": "HKD/升",
                    "value": r["retail"],
                    "diff": None,
                })
                break

    if intl:
        for r in intl["rows"]:
            items.append({
                "label": r["name"],
                "unit": "美元/桶",
                "value": r["cur"],
                "diff": r["diff"],
            })

    if vn:
        for r in vn["rows"]:
            if r["prefix"] == "DO005S_M2":
                items.append({
                    "label": "越南柴油 0.05S（区域 1）",
                    "unit": "VND/升",
                    "value": r["z1"],
                    "diff": None,
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
            diff_str = f"<span style='color:{color};'>{fmt_signed(it['diff'])}</span>"
        val = it["value"]
        val_str = fmt_num(val) if it["unit"] != "VND/升" else fmt_int(val)
        html += (
            f"<tr>"
            f"<td style='color:#5a6b7b;'>{it['label']}</td>"
            f"<td align='right' style='font-weight:600;'>"
            f"{val_str} <span style='color:#8492a6;font-weight:400;'>{it['unit']}</span></td>"
            f"<td align='right' style='width:120px;'>较上次：{diff_str}</td>"
            f"</tr>"
        )
    html += "</table></div>"
    return html

# ==================== 明细 ====================
def section(title, content):
    return f"""
    <div style="margin:24px 0;">
      <h3 style="color:#c9302c;border-left:4px solid #c9302c;
                 padding-left:10px;margin:0 0 10px;font-size:15px;">{title}</h3>
      {content}
    </div>
    """

def render_domestic(rows):
    if not rows:
        return "<p style='color:#e74c3c'>国内数据读取失败</p>"
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
        <tr><th align="left">油品</th><th>代表品牌</th><th>零售价 (HKD/L)</th></tr>
      </thead><tbody>
    """
    for r in d["rows"]:
        html += (
            f"<tr><td>{r['fuel']}</td><td align='center'>{r['brand']}</td>"
            f"<td align='right'>{fmt_num(r['retail'])}</td></tr>"
        )
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
                f"{fmt_signed(r['diff'])}（{fmt_signed(r['pct'])}%）</span>"
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
def build_email():
    today = bj_now().strftime("%Y-%m-%d")

    domestic = fetch_domestic()
    hk = fetch_hongkong()
    intl = fetch_international()
    vn = fetch_vietnam()

    summary_html = build_summary(domestic, hk, intl, vn)

    body = f"""
    <html><body style="font-family:Arial,'Microsoft YaHei',sans-serif;
                        color:#1f2d3d;max-width:960px;margin:auto;padding:16px;">
      <h2 style="color:#c9302c;text-align:center;">每日油价汇总 · {today}</h2>
      <p style="color:#8492a6;font-size:12px;text-align:center;">
        自动推送，数据仅供参考
      </p>

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
    return today, body

# ==================== 发送 ====================
def main():
    if not MAIL_USERNAME or not MAIL_PASSWORD:
        raise SystemExit("缺少 MAIL_USERNAME 或 MAIL_PASSWORD 环境变量")

    today, html = build_email()

    msg = MIMEMultipart("alternative")
    msg["Subject"] = Header(f"【每日油价】{today}", "utf-8")
    msg["From"] = MAIL_USERNAME
    msg["To"] = RECEIVER_EMAIL
    msg.attach(MIMEText(html, "html", "utf-8"))

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
            server.login(MAIL_USERNAME, MAIL_PASSWORD)
            server.sendmail(MAIL_USERNAME, [RECEIVER_EMAIL], msg.as_string())
        print("邮件发送成功 →", RECEIVER_EMAIL)
    except Exception as e:
        print("邮件发送失败:", e)
        raise

if __name__ == "__main__":
    main()
