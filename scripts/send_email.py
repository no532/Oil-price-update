#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日油价推送：
  国内各省市 + 香港 + 国际原油 + 越南
通过 Resend 发送
"""

import os
import json
import requests
import resend
from datetime import datetime, timedelta, timezone

# ==================== 配置 ====================
RESEND_API_KEY = os.environ.get("RESEND_API_KEY")
RECEIVER_EMAIL = "1903859410@qq.com"       # ★ 收件邮箱
SENDER_EMAIL   = "onboarding@resend.dev"              # Resend 测试发件地址

# GitHub Pages 域名
BASE_URL = "https://no532.github.io/Oil-price-update"

# 本地根目录（脚本在 scripts/ 下，往上一级是仓库根目录）
LOCAL_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ==================== 时间 ====================
def bj_now():
    return datetime.now(timezone(timedelta(hours=8)))

# ==================== 通用工具 ====================
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

def load_json(prefer_url, prefer_local):
    d = http_get_json(prefer_url)
    if d is not None:
        return d
    return load_json_local(prefer_local)

# ==================== 1. 国内各省市 ====================
def fetch_domestic():
    """
    读 国内/data.json
    结构：{ "北京市": [ {date, 92号汽油, 92号汽油_diff, ...}, ... ], "天津市": [...], ... }
    取每个省的**最新一条**记录。
    """
    url = f"{BASE_URL}/国内/data.json"
    local = os.path.join(LOCAL_BASE, "国内", "data.json")
    d = load_json(url, local)
    if not d:
        return "<p style='color:#e74c3c'>国内数据读取失败</p>"

    # 每个省取第一条（最新的）
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
            "d92": latest.get("92号汽油_diff"),
            "d95": latest.get("95号汽油_diff"),
            "d0":  latest.get("0号柴油_diff"),
        })

    # 按省份名排序
    rows.sort(key=lambda x: x["province"])

    html = f"""
    <p style="color:#8492a6;font-size:12px;">
      共 {len(rows)} 个省市，显示每省最新一条记录
    </p>
    <table border="1" cellpadding="8" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;width:100%;">
      <thead style="background:#fafbfc;">
        <tr>
          <th align="left">省份</th>
          <th>日期</th>
          <th>92# (元/升)</th>
          <th>95# (元/升)</th>
          <th>0# 柴油 (元/升)</th>
        </tr>
      </thead>
      <tbody>
    """
    for r in rows:
        html += (
            f"<tr>"
            f"<td>{r['province']}</td>"
            f"<td align='center'>{r['date']}</td>"
            f"<td align='right'>{fmt_num(r['p92'])}</td>"
            f"<td align='right'>{fmt_num(r['p95'])}</td>"
            f"<td align='right'>{fmt_num(r['p0'])}</td>"
            f"</tr>"
        )
    html += "</tbody></table>"
    return html

# ==================== 2. 香港 ====================
def fetch_hongkong():
    """
    读 香港/hk_fuel.json
    结构：{ date, diesel: {品牌: {retail, after_discount, discount}}, avg_today, avg_today_retail, open_data: {...} }
    """
    url = f"{BASE_URL}/香港/hk_fuel.json"
    local = os.path.join(LOCAL_BASE, "香港", "hk_fuel.json")
    d = load_json(url, local)
    if not d:
        return "<p style='color:#e74c3c'>香港数据读取失败</p>"

    html = f"""
    <p style="color:#8492a6;font-size:12px;">
      数据日期：{d.get('date','—')}　更新时间：{d.get('update_time','—')}
    </p>
    <table border="1" cellpadding="8" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;width:100%;">
      <thead style="background:#fafbfc;">
        <tr>
          <th align="left">油品</th>
          <th>零售价 (HKD/L)</th>
          <th>优惠后 (HKD/L)</th>
        </tr>
      </thead>
      <tbody>
    """
    # open_data 里是每个油品各品牌的零售价，我们取中石化做代表
    open_data = d.get("open_data", {})
    for fuel_name, brands in open_data.items():
        if not isinstance(brands, dict):
            continue
        # 取中石化作为示例（没有就取第一个品牌）
        sample_brand = "中石化" if "中石化" in brands else list(brands.keys())[0]
        retail = brands.get(sample_brand)
        html += (
            f"<tr>"
            f"<td>{fuel_name}</td>"
            f"<td align='right'>{fmt_num(retail)}</td>"
            f"<td align='right'>—</td>"
            f"</tr>"
        )
    html += "</tbody></table>"
    return html

# ==================== 3. 国际原油 ====================
def fetch_international():
    """
    读 国际/oil_minline.json
    结构：
    {
      "update_time": "...",
      "series": {"WTI美国": [["06:00", 96.45], ...], "布伦特": [...]},
      "meta": {"WTI美国": {"open":..., "prev_close":..., "high":..., "low":...}, ...}
    }
    """
    url = f"{BASE_URL}/国际/oil_minline.json"
    local = os.path.join(LOCAL_BASE, "国际", "oil_minline.json")
    d = load_json(url, local)
    if not d:
        return "<p style='color:#e74c3c'>国际原油数据读取失败</p>"

    html = f"""
    <p style="color:#8492a6;font-size:12px;">更新时间：{d.get('update_time','—')}</p>
    <table border="1" cellpadding="8" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;width:100%;">
      <thead style="background:#fafbfc;">
        <tr>
          <th align="left">品种</th>
          <th>当前 (美元/桶)</th>
          <th>较今开</th>
          <th>今开</th>
          <th>昨收</th>
        </tr>
      </thead>
      <tbody>
    """
    for name, pts in (d.get("series") or {}).items():
        if not pts:
            continue
        cur = pts[-1][1]
        meta = (d.get("meta") or {}).get(name, {})
        open_ = meta.get("open")
        prev  = meta.get("prev_close")

        if open_ and open_ != 0:
            diff = cur - open_
            pct  = diff / open_ * 100
            color = "#e74c3c" if diff > 0 else ("#27ae60" if diff < 0 else "#8492a6")
            sign = "+" if diff > 0 else ""
            pct_str = f"<span style='color:{color}'>{sign}{diff:.2f}（{sign}{pct:.2f}%）</span>"
        else:
            pct_str = "—"

        html += (
            f"<tr>"
            f"<td>{name}</td>"
            f"<td align='right'>{cur:.2f}</td>"
            f"<td align='center'>{pct_str}</td>"
            f"<td align='right'>{fmt_num(open_)}</td>"
            f"<td align='right'>{fmt_num(prev)}</td>"
            f"</tr>"
        )
    html += "</tbody></table>"
    return html

# ==================== 4. 越南 ====================
def fetch_vietnam():
    """
    读 越南/vietnam_fuel.json
    结构：{ date, fuel_names: {prefix: name}, prices: {prefix: {Z1, Z2}} }
    """
    url = f"{BASE_URL}/越南/vietnam_fuel.json"
    local = os.path.join(LOCAL_BASE, "越南", "vietnam_fuel.json")
    d = load_json(url, local)
    if not d:
        return "<p style='color:#e74c3c'>越南数据读取失败</p>"

    order = ["DO005S_M2","DO001S_M5","E10RON95_M5","E10RON95_M3","E5RON92_M2","Kero2K"]
    names = d.get("fuel_names", {})
    prices = d.get("prices", {})

    html = f"""
    <p style="color:#8492a6;font-size:12px;">
      数据日期：{d.get('date','—')}　更新时间：{d.get('update_time','—')}
    </p>
    <table border="1" cellpadding="8" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;width:100%;">
      <thead style="background:#fafbfc;">
        <tr>
          <th align="left">油品</th>
          <th>区域 1 (VND/L)</th>
          <th>区域 2 (VND/L)</th>
        </tr>
      </thead>
      <tbody>
    """
    for k in order:
        p = prices.get(k)
        if not p:
            continue
        nm = names.get(k, k)
        z1 = p.get("Z1")
        z2 = p.get("Z2")
        html += (
            f"<tr>"
            f"<td>{nm}</td>"
            f"<td align='right'>{fmt_int(z1)}</td>"
            f"<td align='right'>{fmt_int(z2)}</td>"
            f"</tr>"
        )
    html += "</tbody></table>"
    return html

# ==================== 数值格式化 ====================
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

# ==================== 拼邮件 ====================
def build_email_html():
    today = bj_now().strftime("%Y-%m-%d")

    def section(title, color, content):
        return f"""
        <div style="margin:24px 0;">
          <h3 style="color:{color};border-left:4px solid {color};
                     padding-left:10px;margin:0 0 10px;">{title}</h3>
          {content}
        </div>
        """

    body = f"""
    <html><body style="font-family:Arial,'Microsoft YaHei',sans-serif;
                        color:#1f2d3d;max-width:960px;margin:auto;padding:16px;">
      <h2 style="color:#c9302c;text-align:center;">每日油价汇总 · {today}</h2>
      <p style="color:#8492a6;font-size:12px;text-align:center;">
        自动推送，数据仅供参考
      </p>

      {section("一、国内各省市油价", "#c9302c", fetch_domestic())}
      {section("二、香港油价",       "#c9302c", fetch_hongkong())}
      {section("三、国际原油",       "#c9302c", fetch_international())}
      {section("四、越南油价",       "#c9302c", fetch_vietnam())}

      <hr style="border:none;border-top:1px dashed #e4e7ed;margin:24px 0;">
      <p style="color:#8492a6;font-size:12px;text-align:center;">
        本邮件由 GitHub Actions 自动生成，请勿回复<br>
        数据来源：<a href="{BASE_URL}" style="color:#409eff;">{BASE_URL}</a>
      </p>
    </body></html>
    """
    return today, body

# ==================== 发送 ====================
def main():
    if not RESEND_API_KEY:
        raise SystemExit("缺少 RESEND_API_KEY 环境变量")

    resend.api_key = RESEND_API_KEY

    today, html = build_email_html()

    params = {
        "from": SENDER_EMAIL,
        "to": [RECEIVER_EMAIL],
        "subject": f"【每日油价】{today}",
        "html": html,
    }

    try:
        result = resend.Emails.send(params)
        print("邮件发送成功:", result)
    except Exception as e:
        print("邮件发送失败:", e)
        raise

if __name__ == "__main__":
    main()
