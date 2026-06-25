#!/usr/bin/env python3
"""
建筑汇报 PPT 自动助手 — HTML 多页预览生成器
根据结构化数据生成可预览的 HTML 演示页面。
"""

import json
import argparse
import os
from datetime import datetime

DEFAULT_SECTIONS = {
    "title": "某新城核心区城市设计",
    "subtitle": "概念规划方案汇报",
    "client": "某市自然资源和规划局",
    "date": datetime.now().strftime("%Y年%m月"),
    "pages": [
        {
            "type": "cover",
            "title": "某新城核心区城市设计",
            "subtitle": "概念规划方案汇报",
            "meta": "某市自然资源和规划局 · 2025"
        },
        {
            "type": "toc",
            "title": "汇报框架",
            "items": [
                "01 / 项目区位",
                "02 / 现状分析",
                "03 / 设计策略",
                "04 / 总图方案",
                "05 / 建筑分析",
                "06 / 效果展示"
            ]
        },
        {
            "type": "content",
            "title": "01 / 项目区位",
            "subtitle": "基地位于新城核心区，交通枢纽辐射范围 3km",
            "body": "项目位于新城核心区东侧，紧邻城市主干道及地铁3号线站点，周边配套完善。基地总面积约12.6公顷，现状以空地及低效工业用地为主。",
            "stats": [
                {"label": "基地面积", "value": "12.6 ha"},
                {"label": "容积率", "value": "2.8"},
                {"label": "限高", "value": "80 m"},
                {"label": "绿地率", "value": "≥ 35%"}
            ]
        },
        {
            "type": "content",
            "title": "02 / 现状分析",
            "subtitle": "周边交通、景观资源与用地条件分析",
            "body": "基地南侧为城市主干道，东侧规划地铁站点，北侧为滨河景观带。现状以空地为主，局部有临时厂房需拆改。周边已建成高品质住宅、商业综合体及学校，城市界面良好。",
            "stats": [
                {"label": "主干道", "value": "60 m"},
                {"label": "地铁站", "value": "500 m"},
                {"label": "滨河带", "value": "300 m"},
                {"label": "现状建筑", "value": "1.2 万㎡"}
            ]
        },
        {
            "type": "split",
            "title": "03 / 设计策略",
            "left": "策略一：TOD 导向开发，以地铁站点为核心，形成 TOD 综合开发区，提高土地价值。",
            "right": "策略二：生态优先，保留滨河绿带并延伸至基地内部，构建蓝绿网络体系。"
        },
        {
            "type": "content",
            "title": "04 / 总图方案",
            "subtitle": "一轴两带三组团 · 功能复合的城市核心",
            "body": "总图以 \"一轴两带三组团\" 为结构：南北向城市活力轴串联各功能组团；东西两条生态绿带渗透基地；形成 TOD 综合组团、品质居住组团、滨水商业组团三大板块。",
            "stats": [
                {"label": "总建筑面积", "value": "35.2 万㎡"},
                {"label": "商业", "value": "8.6 万㎡"},
                {"label": "住宅", "value": "18.4 万㎡"},
                {"label": "公建配套", "value": "8.2 万㎡"}
            ]
        },
        {
            "type": "content",
            "title": "05 / 建筑分析",
            "subtitle": "高度分区与天际线控制",
            "body": "建筑高度呈梯度分布：沿河及绿地周边以多层为主（≤24m），核心区布局高层（60-80m），形成\"外低内高、逐级递进\"的天际线节奏。裙房采用暖色石材+玻璃幕墙，塔楼以浅色铝板为主。",
            "stats": [
                {"label": "多层区", "value": "≤24 m"},
                {"label": "高层区", "value": "60-80 m"},
                {"label": "塔楼数量", "value": "6 栋"},
                {"label": "立面材质", "value": "铝板+玻璃"}
            ]
        },
        {
            "type": "end",
            "title": "06 / 效果展示",
            "subtitle": "感谢聆听 · 敬请指正"
        }
    ]
}


def generate_html(data):
    pages_html = ""
    for i, page in enumerate(data["pages"]):
        pages_html += _render_page(page, i)

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{data['title']} · 预览</title>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap" rel="stylesheet">
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    html {{ scroll-behavior: smooth; scroll-snap-type: y proximity; }}
    body {{
      font-family: 'Inter', system-ui, -apple-system, sans-serif;
      background: #141413;
      color: #faf9f5;
      -webkit-font-smoothing: antialiased;
    }}

    /* Layout — mimics A4 portrait at comfortable height */
    .slide {{
      min-height: 100vh;
      padding: 4rem 3rem;
      display: flex;
      flex-direction: column;
      justify-content: center;
      scroll-snap-align: start;
      border-bottom: 1px solid #30302e;
      max-width: 960px;
      margin: 0 auto;
    }}
    .slide:last-child {{ border-bottom: none; }}

    /* Typography */
    .slide-label {{
      font-size: 0.7rem;
      font-weight: 500;
      letter-spacing: 0.12px;
      text-transform: uppercase;
      color: #87867f;
      margin-bottom: 0.5rem;
    }}
    .slide h1 {{
      font-family: Georgia, 'Times New Roman', Times, serif;
      font-size: clamp(2rem, 4vw, 3.5rem);
      font-weight: 500;
      line-height: 1.15;
      color: #faf9f5;
      margin-bottom: 0.75rem;
    }}
    .slide h2 {{
      font-family: Georgia, 'Times New Roman', Times, serif;
      font-size: clamp(1.4rem, 2.5vw, 2rem);
      font-weight: 500;
      line-height: 1.25;
      color: #faf9f5;
      margin-bottom: 0.75rem;
    }}
    .slide h3 {{
      font-family: Georgia, 'Times New Roman', Times, serif;
      font-size: 1.15rem;
      font-weight: 500;
      color: #d97757;
      margin-bottom: 0.5rem;
    }}
    .slide p {{
      color: #b0aea5;
      line-height: 1.8;
      font-size: 0.95rem;
      max-width: 640px;
    }}
    .slide .subtitle {{
      color: #87867f;
      font-size: 1rem;
      line-height: 1.6;
      margin-bottom: 1.5rem;
      max-width: 600px;
    }}

    /* Cover */
    .cover {{ text-align: center; }}
    .cover .meta {{ color: #87867f; font-size: 0.85rem; margin-top: 1.5rem; }}
    .cover .accent-line {{
      width: 40px; height: 3px; background: #c96442;
      margin: 1.5rem auto; border-radius: 2px;
    }}

    /* TOC */
    .toc-list {{ list-style: none; display: flex; flex-direction: column; gap: 0.6rem; }}
    .toc-list li {{
      font-size: 1.05rem; color: #b0aea5;
      transition: color 0.2s; cursor: default;
    }}
    .toc-list li:hover {{ color: #faf9f5; }}

    /* Stats Grid */
    .stat-grid {{
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 1px;
      background: #30302e;
      border-radius: 10px;
      overflow: hidden;
      margin-top: 2rem;
      max-width: 600px;
    }}
    .stat-grid .stat {{ background: #1a1a19; padding: 1.25rem 1rem; text-align: center; }}
    .stat .val {{ font-size: 1.15rem; font-weight: 500; color: #faf9f5; }}
    .stat .lbl {{ font-size: 0.7rem; color: #87867f; margin-top: 0.15rem; }}

    /* Split */
    .split-grid {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 2rem;
      margin-top: 1.5rem;
    }}
    .split-card {{
      background: #1a1a19;
      border: 1px solid #30302e;
      border-radius: 12px;
      padding: 2rem;
    }}
    .split-card p {{ font-size: 0.95rem; color: #b0aea5; line-height: 1.7; }}

    /* End slide */
    .end {{ text-align: center; justify-content: center; align-items: center; }}

    /* Nav dots */
    .nav-bar {{
      position: fixed; bottom: 2rem; right: 2rem;
      display: flex; gap: 0.4rem; z-index: 100;
    }}
    .nav-dot {{
      width: 8px; height: 8px; border-radius: 50%;
      background: #30302e; border: none; cursor: pointer;
      transition: background 0.2s;
    }}
    .nav-dot:hover {{ background: #87867f; }}

    /* Toolbar */
    .toolbar {{
      position: fixed; top: 1rem; right: 1rem; z-index: 100;
      display: flex; gap: 0.5rem;
    }}
    .toolbar a {{
      display: inline-block;
      background: #1a1a19; border: 1px solid #30302e;
      color: #b0aea5; padding: 0.4rem 0.8rem;
      border-radius: 6px; text-decoration: none;
      font-size: 0.75rem; transition: all 0.2s;
    }}
    .toolbar a:hover {{ background: #30302e; color: #faf9f5; }}

    @media (max-width: 640px) {{
      .slide {{ padding: 2rem 1.5rem; }}
      .stat-grid {{ grid-template-columns: repeat(2, 1fr); }}
      .split-grid {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>

<div class="toolbar">
  <a href="#" onclick="window.print()">打印 / 导出 PDF</a>
  <a href="#" onclick="document.querySelector('.slide').scrollIntoView({{behavior:'smooth'}})">回到顶部</a>
</div>

<div class="nav-bar" id="navDots"></div>

{''.join(pages_html)}

<script>
  // Generate nav dots
  const slides = document.querySelectorAll('.slide');
  const nav = document.getElementById('navDots');
  slides.forEach((_, i) => {{
    const dot = document.createElement('button');
    dot.className = 'nav-dot';
    dot.addEventListener('click', () => slides[i].scrollIntoView({{ behavior: 'smooth' }}));
    nav.appendChild(dot);
  }});

  // Highlight current dot
  const dots = nav.querySelectorAll('.nav-dot');
  const observer = new IntersectionObserver((entries) => {{
    entries.forEach(entry => {{
      if (entry.isIntersecting) {{
        const idx = Array.from(slides).indexOf(entry.target);
        dots.forEach((d, i) => d.style.background = i === idx ? '#c96442' : '#30302e');
      }}
    }});
  }}, {{ threshold: 0.6 }});
  slides.forEach(s => observer.observe(s));
</script>

</body>
</html>"""


def _render_page(page, idx):
    t = page["type"]
    if t == "cover":
        return f"""<div class="slide cover">
  <div class="slide-label">建筑汇报方案</div>
  <h1>{page['title']}</h1>
  <div class="accent-line"></div>
  <p style="font-size:1.15rem; color:#87867f;">{page['subtitle']}</p>
  <div class="meta">{page['meta']}</div>
</div>"""

    if t == "toc":
        items = "".join(f"<li>{item}</li>" for item in page["items"])
        return f"""<div class="slide">
  <div class="slide-label">汇报框架</div>
  <h2>{page['title']}</h2>
  <ul class="toc-list" style="margin-top:2rem;">{items}</ul>
</div>"""

    if t == "content":
        stats_html = ""
        if page.get("stats"):
            stats = "".join(
                f'<div class="stat"><div class="val">{s["value"]}</div><div class="lbl">{s["label"]}</div></div>'
                for s in page["stats"]
            )
            stats_html = f'<div class="stat-grid">{stats}</div>'

        return f"""<div class="slide">
  <div class="slide-label">第 {idx+1} 页</div>
  <h2>{page['title']}</h2>
  <p class="subtitle">{page['subtitle']}</p>
  <p>{page['body']}</p>
  {stats_html}
</div>"""

    if t == "split":
        return f"""<div class="slide">
  <div class="slide-label">第 {idx+1} 页</div>
  <h2>{page['title']}</h2>
  <div class="split-grid">
    <div class="split-card"><h3>策略一</h3><p>{page['left']}</p></div>
    <div class="split-card"><h3>策略二</h3><p>{page['right']}</p></div>
  </div>
</div>"""

    if t == "end":
        return f"""<div class="slide end">
  <div class="accent-line" style="margin-bottom:1.5rem;"></div>
  <h2>{page['title']}</h2>
  <p class="subtitle" style="margin:0 auto;">{page['subtitle']}</p>
</div>"""

    return ""


def main():
    parser = argparse.ArgumentParser(description="建筑汇报 HTML 预览生成器")
    parser.add_argument("--out", "-o", default="preview.html", help="输出文件路径")
    parser.add_argument("--json", "-j", help="从 JSON 文件读取页面数据")
    args = parser.parse_args()

    if args.json:
        with open(args.json, "r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        data = DEFAULT_SECTIONS

    html = generate_html(data)
    out_path = os.path.abspath(args.out)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ 预览已生成: {out_path}")
    print(f"   共 {len(data['pages'])} 页 · 项目: {data['title']}")


if __name__ == "__main__":
    main()
