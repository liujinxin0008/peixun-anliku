"""
人力看板数据加工脚本
数据来源：Step1 人力明细 & Step5 人员变动明细表（全量月份）
输出：data.json + talent_dashboard.html（数据直接嵌入，双击浏览器可直接打开）
"""
import pandas as pd
import json
import os
import re
import glob

SNAPSHOT_DIR = r"e:\Claude Code\人才评估系统Agent\Step1-人力数据加工\outputs"
CHANGES_DIR  = r"e:\Claude Code\人才评估系统Agent\Step5-月度数据产出\人员变动明细表\outputs"
OUTPUT_DIR   = os.path.dirname(os.path.abspath(__file__))
HTML_FILE    = os.path.join(OUTPUT_DIR, "hr_dashboard.html")
JSON_FILE    = os.path.join(OUTPUT_DIR, "data.json")


def load_snapshots():
    files = sorted(f for f in glob.glob(os.path.join(SNAPSHOT_DIR, "*人力明细.xlsx"))
                   if not os.path.basename(f).startswith("~$"))
    if not files:
        raise FileNotFoundError(f"未找到人力明细文件：{SNAPSHOT_DIR}")
    records = []
    for f in files:
        df = pd.read_excel(f)
        # 本月标签优先，全空时 fallback 到人员状态
        if "本月标签" in df.columns and df["本月标签"].notna().any():
            status_src = "本月标签"
        else:
            status_src = "人员状态"
        df = df.rename(columns={"数据月份": "month", "城市": "city", "工号": "employee_id",
                                 "姓名": "name", status_src: "status", "入职时间": "join_date",
                                 "晋升时间": "promotion_date"})
        keep = [c for c in ["month", "city", "employee_id", "name", "status", "join_date", "promotion_date"] if c in df.columns]
        df = df[keep].copy()
        # 只保留主任和见习主任
        df = df[df["status"].isin(["主任", "见习主任", "协催主任"])]
        df["month"] = df["month"].astype(str).str[:7]
        df["join_date"] = pd.to_datetime(df["join_date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
        df["promotion_date"] = pd.to_datetime(df["promotion_date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
        records.extend(df.to_dict(orient="records"))
    print(f"快照记录：{len(records)} 条，来自 {len(files)} 个文件")
    return records


def load_changes(snapshots):
    files = sorted(f for f in glob.glob(os.path.join(CHANGES_DIR, "*人员变动明细表.xlsx"))
                   if not os.path.basename(f).startswith("~$"))
    if not files:
        raise FileNotFoundError(f"未找到人员变动明细表文件：{CHANGES_DIR}")
    city_map = {s["employee_id"]: s["city"] for s in snapshots if s.get("city")}
    records = []
    for f in files:
        df = pd.read_excel(f)
        if df.empty or len(df.columns) == 0:
            continue
        df = df.rename(columns={"变动月份": "month", "工号": "employee_id", "姓名": "name",
                                 "变动类型": "change_type", "备注": "remark"})
        keep = [c for c in ["month", "employee_id", "name", "change_type", "remark"] if c in df.columns]
        df = df[keep].copy()
        if "month" not in df.columns:
            continue
        df["month"] = df["month"].astype(str).str[:7]
        df["remark"] = df["remark"].fillna("").astype(str)
        df["city"] = df["employee_id"].map(city_map).fillna("未知")
        records.extend(df.to_dict(orient="records"))
    print(f"变动记录：{len(records)} 条，来自 {len(files)} 个文件")
    return records


def inject_data_into_html(data):
    """读取 HTML，将数据嵌入，使看板无需 fetch 即可直接打开"""
    with open(HTML_FILE, "r", encoding="utf-8") as f:
        html = f.read()

    # 1. 清除上次注入的数据块（幂等）
    html = re.sub(r'\n?<script id="inline-data">.*?</script>', "", html, flags=re.DOTALL)

    # 2. 在 </head> 前注入新数据
    data_script = f'\n<script id="inline-data">const __DATA__ = {json.dumps(data, ensure_ascii=False)};</script>'
    html = html.replace("</head>", data_script + "\n</head>", 1)

    # 3. 替换 loadData 中的 fetch 逻辑为内联数据读取（幂等：先恢复原始再替换）
    # 先确保是原始 fetch 版本（兼容多次运行）
    inline_block = (
        "const json = __DATA__; // 数据已内嵌，无需 fetch"
    )
    fetch_block = (
        "const response = await fetch('./data.json'); // 数据文件路径\n"
        "            if (!response.ok) throw new Error('数据加载失败');\n"
        "            const json = await response.json();"
    )
    inline_marker = "const json = __DATA__; // 数据已内嵌，无需 fetch"

    if inline_marker in html:
        # 已替换过，不需要再替换
        pass
    elif fetch_block in html:
        html = html.replace(fetch_block, inline_block, 1)

    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"hr_dashboard.html 已更新（数据内嵌）：{HTML_FILE}")


def main():
    print("开始数据加工...")
    snapshots = load_snapshots()
    changes   = load_changes(snapshots)
    data = {"snapshots": snapshots, "changes": changes}

    # 输出 data.json
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"data.json 已输出：{JSON_FILE}")

    # 将数据嵌入 HTML
    inject_data_into_html(data)
    print("完成！")


if __name__ == "__main__":
    main()
