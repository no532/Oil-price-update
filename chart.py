import csv
import json
from collections import defaultdict

CSV_FILE = "oil.csv"
OUT_HTML = "oil_chart.html"

def load():
    series = defaultdict(list)  # 品种 -> [(时间, 价格)]
    with open(CSV_FILE, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if not row["当前价"]:
                continue
            series[row["品种"]].append((row["抓取时间"], float(row["当前价"])))
    return series

def build_html(series):
    data_js = json.dumps(
        {k: {"x": [t for t, _ in v], "y": [p for _, p in v]} for k, v in series.items()},
        ensure_ascii=False
    )
    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>国际原油走势</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"></script>
<style>
body {{ font-family: -apple-system,"Microsoft YaHei",sans-serif; margin:0; padding:24px; background:#f0f2f5; }}
.card {{ background:#fff; border-radius:12px; padding:20px; box-shadow:0 2px 12px rgba(0,0,0,.06); }}
h2 {{ margin:0 0 16px; font-size:18px; }}
#chart {{ width:100%; height:480px; }}
</style></head>
<body>
<div class="card">
  <h2>国际原油走势（数据源：新浪财经）</h2>
  <div id="chart"></div>
</div>
<script>
const DATA = {data_js};
const chart = echarts.init(document.getElementById('chart'));
const colors = ['#e74c3c', '#2c5fa8', '#16a085'];
const series = Object.keys(DATA).map((name, i) => ({{
  name, type:'line', smooth:true, symbol:'circle', symbolSize:5,
  itemStyle:{{ color: colors[i % colors.length] }},
  data: DATA[name].y
}}));
const xAxis = Object.values(DATA)[0]?.x || [];
chart.setOption({{
  tooltip: {{ trigger:'axis' }},
  legend: {{ top:0 }},
  grid: {{ left:60, right:30, top:50, bottom:80 }},
  xAxis: {{ type:'category', data:xAxis, axisLabel:{{ rotate:45, fontSize:10 }} }},
  yAxis: {{ type:'value', name:'美元/桶', scale:true }},
  series
}});
</script>
</body></html>"""
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"已生成 {OUT_HTML}")

if __name__ == "__main__":
    build_html(load())