import pandas as pd
import json
import os
import glob

# ==================== 配置路径 ====================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MAIN_DATA_DIR = r"E:\Claude Code\人才评估系统Agent\Step5-月度数据产出\历史事实表\outputs"
ABILITY_FILE = r"E:\Claude Code\人才评估系统Agent\Step6-人才评估数据加工\input\主任能力评分.xlsx"
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "employee_profile.js")

# ==================== 读取最新历史事实表（已含历史累积数据）====================
all_files = sorted([
    f for f in glob.glob(os.path.join(MAIN_DATA_DIR, "*-历史事实表.xlsx"))
    if not os.path.basename(f).startswith("~$")
])
latest_file = all_files[-1]
print(f"使用最新历史事实表：{os.path.basename(latest_file)}")

df_main = pd.read_excel(latest_file)

# 读取能力数据
df_ability = pd.read_excel(ABILITY_FILE)

# ==================== 数据清洗 ====================
df_main["数据月份"] = pd.to_datetime(df_main["数据月份"])

# 能力数据月份：202602 -> 2026-02
df_ability["月份"] = pd.to_datetime(df_ability["月份"].astype(str), format="%Y%m")

# 百分比字符串 -> float（兼容 "85.5%" 和 0.855 两种格式）
def parse_pct(val):
    if pd.isna(val):
        return float("nan")
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace("%", "")
    try:
        v = float(s)
        # 如果原始带 % 符号，已除以 100；否则判断量级
        return v / 100 if "%" in str(val) else v
    except Exception:
        return float("nan")

for col in ["小组队列排名占比", "小组目标达成率", "话术目标达成率",
            "当月员工离职率", "当月作业完成率", "当月培训出勤率"]:
    if col in df_main.columns:
        df_main[col] = df_main[col].apply(parse_pct)

# 其他数值字段强制转换
for col in ["话术平均分", "早会/夕会得分", "PIP面谈得分", "离职面谈得分", "人员辅导得分", "管理过程总得分"]:
    df_main[col] = pd.to_numeric(df_main[col], errors="coerce")

# 同一员工同月份去重（保留最后一条）
df_main = df_main.drop_duplicates(subset=["工号", "数据月份"], keep="first")

# ==================== 最新月份 ====================
latest_month_main = df_main["数据月份"].max()
latest_month_ability = df_ability["月份"].max()

print(f"主数据最新月份: {latest_month_main.strftime('%Y-%m')}")
print(f"能力数据最新月份: {latest_month_ability.strftime('%Y-%m')}")

df_main_latest = df_main[df_main["数据月份"] == latest_month_main]
df_ability_latest = df_ability[df_ability["月份"] == latest_month_ability]

# 团队均值（按城市，基于最新能力数据）
team_avg_dict = (
    df_ability_latest.groupby("城市")[
        ["角色认知与目标规划", "过程管控与复盘", "绩效推进与沟通",
         "团队激励与士气", "人才识别与授权", "员工辅导与发展"]
    ].mean()
    .to_dict(orient="index")
)

# ==================== 工具函数 ====================
def sf(val):
    """safe float"""
    try:
        v = float(val)
        return 0.0 if pd.isna(v) else v
    except Exception:
        return 0.0

def si(val):
    """safe int"""
    try:
        v = float(val)
        return 0 if pd.isna(v) else int(v)
    except Exception:
        return 0

def make_mgmt_trend(recent_group, col_count, col_score):
    trend = []
    for _, row in recent_group.iterrows():
        count = si(row[col_count])
        score = round(sf(row[col_score]), 2) if count > 0 else 0.0
        trend.append({"month": row["数据月份"].strftime("%Y-%m"), "count": count, "avg_score": score})
    return trend

def make_value_trend(recent_group, col, decimals=4):
    return [
        {"month": row["数据月份"].strftime("%Y-%m"), "value": round(sf(row[col]), decimals)}
        for _, row in recent_group.iterrows()
    ]

def make_int_trend(recent_group, col):
    return [
        {"month": row["数据月份"].strftime("%Y-%m"), "value": si(row[col])}
        for _, row in recent_group.iterrows()
    ]

def make_rank_trend(recent_group):
    """排名占比趋势，附带 rank_detail（如 '1/4'）"""
    trend = []
    for _, row in recent_group.iterrows():
        val = round(sf(row["小组队列排名占比"]), 4)
        detail = ""
        if "排名占比明细" in row.index:
            detail = str(row["排名占比明细"] or "").strip()
            if detail.lower() == "nan":
                detail = ""
        trend.append({
            "month": row["数据月份"].strftime("%Y-%m"),
            "value": val,
            "rank_detail": detail,
        })
    return trend

def avg_safe(series):
    vals = [sf(v) for v in series]
    return sum(vals) / len(vals) if vals else 0.0

def avg_nonzero(series):
    """排除0值计算平均"""
    vals = [sf(v) for v in series if sf(v) > 0]
    return sum(vals) / len(vals) if vals else 0.0

# ==================== 生成员工档案 ====================
result = []

for _, latest in df_main_latest.iterrows():
    emp_id = latest["工号"]
    city = latest["城市"]
    status = latest["本月标签"] if pd.notna(latest["本月标签"]) and latest["本月标签"] != '' else latest["人员状态"]

    if status not in ["主任", "见习主任", "协催主任"]:
        continue

    # 该员工全部历史（用于趋势 & 累计）
    group = df_main[df_main["工号"] == emp_id].sort_values("数据月份")
    recent = group.tail(6)

    # 能力得分（最新月份）
    ab_row = df_ability_latest[df_ability_latest["工号"] == emp_id]
    if not ab_row.empty:
        r = ab_row.iloc[0]
        ability_self = {
            "role_goal": float(r["角色认知与目标规划"]),
            "process_review": float(r["过程管控与复盘"]),
            "performance_comm": float(r["绩效推进与沟通"]),
            "team_morale": float(r["团队激励与士气"]),
            "talent_auth": float(r["人才识别与授权"]),
            "staff_dev": float(r["员工辅导与发展"]),
        }
    else:
        ability_self = {k: 0.0 for k in
                        ["role_goal", "process_review", "performance_comm",
                         "team_morale", "talent_auth", "staff_dev"]}

    # 主管评语（来自能力评分表综合评语列）
    supervisor_comment = ""
    if not ab_row.empty:
        comment_col = df_ability.columns[20]  # 综合评语
        raw = ab_row.iloc[0][comment_col]
        supervisor_comment = str(raw) if pd.notna(raw) else ""

    ts = team_avg_dict.get(city, {})
    ability_team = {
        "role_goal": round(float(ts.get("角色认知与目标规划", 0)), 2),
        "process_review": round(float(ts.get("过程管控与复盘", 0)), 2),
        "performance_comm": round(float(ts.get("绩效推进与沟通", 0)), 2),
        "team_morale": round(float(ts.get("团队激励与士气", 0)), 2),
        "talent_auth": round(float(ts.get("人才识别与授权", 0)), 2),
        "staff_dev": round(float(ts.get("员工辅导与发展", 0)), 2),
    }

    # 基本信息
    try:
        hire_date = pd.to_datetime(latest["入职时间"])
        hire_date_str = hire_date.strftime("%Y-%m-%d")
    except Exception:
        hire_date_str = ""

    try:
        promotion_date = pd.to_datetime(latest["晋升时间"])
        end_of_month = latest_month_main + pd.offsets.MonthEnd(0)
        tenure_months = (end_of_month.year - promotion_date.year) * 12 + \
                        (end_of_month.month - promotion_date.month)
        tenure_months = max(tenure_months, 0)
    except Exception:
        tenure_months = 0

    # 最近一次培训时间
    last_tr = group["最近一次培训时间"].max()
    try:
        last_tr_str = pd.to_datetime(last_tr).strftime("%Y-%m-%d") if not pd.isna(last_tr) else None
    except Exception:
        last_tr_str = None

    # 管理过程趋势
    total_score_trend = [
        {"month": row["数据月份"].strftime("%Y-%m"), "score": round(sf(row["管理过程总得分"]), 2)}
        for _, row in recent.iterrows()
    ]

    emp_json = {
        "employee_id": emp_id,
        "name": latest["姓名"],
        "city": city,
        "status": status,
        "queue": str(latest.get("本月队列", "") or ""),
        "reporter": str(latest.get("汇报上级", "") or ""),
        "hire_date": hire_date_str,
        "tenure_months": tenure_months,
        "supervisor_comment": supervisor_comment,
        "ability": {"self": ability_self, "team_avg": ability_team},
        "training": {
            "total_hours": si(group["当月培训课时"].sum()),
            "avg_score": round(sf(group["当月培训考核平均分"].mean()), 2),
            "last_training_date": last_tr_str,
        },
        "management": {
            "morning_evening": {"trend": make_mgmt_trend(recent, "早会/夕会次数", "早会/夕会得分")},
            "pip":             {"trend": make_mgmt_trend(recent, "PIP面谈次数", "PIP面谈得分")},
            "coaching":        {"trend": make_mgmt_trend(recent, "人员辅导次数", "人员辅导得分")},
            "exit_interview":  {"trend": make_mgmt_trend(recent, "离职面谈次数", "离职面谈得分")},
            "total_score":     {"trend": total_score_trend, "latest": round(sf(latest["管理过程总得分"]), 2)},
        },
        "business": {
            "goal_achievement_rate": {
                "last_3_months_avg": round(avg_safe(group.tail(3)["小组目标达成率"]), 4),
                "trend": make_value_trend(recent, "小组目标达成率", 4),
            },
            "rank_percentile": {
                "avg": round(avg_nonzero(recent["小组队列排名占比"]), 4),
                "trend": make_rank_trend(recent),
            },
            "talk_avg_score": {
                "avg": round(avg_safe(recent["话术平均分"]), 2),
                "trend": make_value_trend(recent, "话术平均分", 2),
            },
            "talk_achievement_rate": {
                "avg": round(avg_safe(recent["话术目标达成率"]), 4),
                "trend": make_value_trend(recent, "话术目标达成率", 4),
            },
        },
        "risk": {
            "complaint": {
                "last_6_months_sum": si(recent["重渠投诉量"].sum()),
                "trend": make_int_trend(recent, "重渠投诉量"),
            },
            "violation": {
                "last_6_months_sum": si(recent["三级及以上质检违规量"].sum()),
                "trend": make_int_trend(recent, "三级及以上质检违规量"),
            },
        },
    }
    result.append(emp_json)

# ==================== 输出 ====================
with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write("var EMPLOYEE_DATA = ")
    json.dump(result, f, ensure_ascii=False, indent=2)
    f.write(";")

print(f"\n[完成] 生成成功：共 {len(result)} 位主任/见习主任")
print(f"[输出] 已保存至：{OUTPUT_FILE}")
