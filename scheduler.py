import time
import subprocess
import sys
from datetime import datetime, timedelta

RUN_HOUR = 9      # 每天几点跑（24小时制）
RUN_MINUTE = 0    # 几分跑
SCRIPT = r"C:\Users\31007384\Desktop\油价自动抓取\国内\测试.py"
LOG    = r"C:\Users\31007384\Desktop\油价自动抓取\国内\scheduler.log"

def log(msg):
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")

def next_run():
    now = datetime.now()
    target = now.replace(hour=RUN_HOUR, minute=RUN_MINUTE, second=0, microsecond=0)
    if target <= now:
        target += timedelta(days=1)
    return target

log("启动 scheduler")
while True:
    t = next_run()
    wait = (t - datetime.now()).total_seconds()
    log(f"下次运行：{t}（还有 {wait/3600:.2f} 小时）")
    time.sleep(wait)
    log("开始运行 测试.py")
    try:
        r = subprocess.run([sys.executable, SCRIPT],
                           capture_output=True, text=True, encoding="utf-8")
        log("输出：\n" + (r.stdout or ""))
        if r.stderr:
            log("错误：\n" + r.stderr)
    except Exception as e:
        log(f"运行失败：{e}")