import pandas as pd
import pulp as pl
from collections import defaultdict
import matplotlib.pyplot as plt

# 设置中文显示
plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False


def load_attachment1(file_path):
    """加载附件1数据"""
    df_tradition = pd.read_excel(file_path, sheet_name="传统电价")
    tradition_price = df_tradition.set_index("时间（时）")[
        "传统电价（单位：元/千瓦时 ）"
    ].to_dict()

    df_new_energy = pd.read_excel(file_path, sheet_name="新能源电价")
    new_energy_price = df_new_energy.set_index("时间（时）")[
        "电价（单位：元/千瓦时 ）"
    ].to_dict()

    df_supply = pd.read_excel(file_path, sheet_name="新能源电力供应量")
    new_energy_supply = (
        df_supply.set_index("时间（时）")["新能源电力供应（兆瓦）"] * 1000
    ).to_dict()

    return tradition_price, new_energy_price, new_energy_supply


def load_attachment2(file_path):
    """加载附件2数据并生成任务结构"""
    df = pd.read_excel(file_path, sheet_name="Sheet1")

    high_tasks = defaultdict(float)
    mid_subtasks = []
    low_subtasks = []

    for _, row in df.iterrows():
        time_range = row.iloc[0]
        high = row.iloc[1]
        mid = row.iloc[2]
        low = row.iloc[3]

        start_str, end_str = time_range.split("-")
        start_hour = int(start_str.split(":")[0])
        end_hour = int(end_str.split(":")[0])
        hours = list(range(start_hour, end_hour))
        num_hours = len(hours)

        if num_hours == 0:
            continue

        per_high = high / num_hours
        for h in hours:
            high_tasks[h] += per_high

        per_mid = mid / num_hours
        per_low = low / num_hours
        for h in hours:
            mid_subtasks.append((per_mid * 50, h))
            low_subtasks.append((per_low * 30, h))

    return high_tasks, mid_subtasks, low_subtasks


def build_and_solve_model(
    tradition_price,
    new_energy_price,
    new_energy_supply,
    high_tasks: defaultdict,
    mid_subtasks,
    low_subtasks,
):
    """构建并求解ILP模型"""
    model = pl.LpProblem("Power_Scheduling_Optimization", pl.LpMinimize)
    hours = range(24)
    beta = 0.15  # 绿色能源使用奖励系数
    gamma = 0.05  # 新能源供应奖励系数

    # 高优先级任务变量
    green_high = {h: pl.LpVariable(f"green_high_{h}", 0) for h in hours}
    trad_high = {h: pl.LpVariable(f"trad_high_{h}", 0) for h in hours}

    # 高优先级任务约束
    for h in hours:
        model += green_high[h] + trad_high[h] == high_tasks.get(h, 0) * 80

    # 中低优先级任务变量存储结构
    mid_green_vars = defaultdict(list)
    mid_trad_vars = defaultdict(list)
    low_green_vars = defaultdict(list)
    low_trad_vars = defaultdict(list)

    # 处理中优先级任务
    for idx, (energy, pub_hour) in enumerate(mid_subtasks):
        allowed_hours = [(pub_hour + i) % 24 for i in range(24)]
        y_vars = {}
        g_vars = {}
        t_vars = {}

        for h in allowed_hours:
            y_var = pl.LpVariable(f"mid_y_{idx}_{h}", cat="Binary")
            g_var = pl.LpVariable(f"mid_g_{idx}_{h}", 0)
            t_var = pl.LpVariable(f"mid_t_{idx}_{h}", 0)

            model += g_var + t_var == energy * y_var
            y_vars[h] = y_var
            g_vars[h] = g_var
            t_vars[h] = t_var

            mid_green_vars[h].append(g_var)
            mid_trad_vars[h].append(t_var)
            model += -gamma * new_energy_supply[h] * y_var  # 供应量奖励

        model += pl.lpSum(y_vars.values()) == 1

    # 处理低优先级任务
    for idx, (energy, pub_hour) in enumerate(low_subtasks):
        allowed_hours = [(pub_hour + i) % 24 for i in range(24)]
        y_vars = {}
        g_vars = {}
        t_vars = {}

        for h in allowed_hours:
            y_var = pl.LpVariable(f"low_y_{idx}_{h}", cat="Binary")
            g_var = pl.LpVariable(f"low_g_{idx}_{h}", 0)
            t_var = pl.LpVariable(f"low_t_{idx}_{h}", 0)

            model += g_var + t_var == energy * y_var
            y_vars[h] = y_var
            g_vars[h] = g_var
            t_vars[h] = t_var

            low_green_vars[h].append(g_var)
            low_trad_vars[h].append(t_var)
            model += -gamma * new_energy_supply[h] * y_var  # 供应量奖励

        model += pl.lpSum(y_vars.values()) == 1

    # 新能源供应约束
    for h in hours:
        total_green = (
            green_high[h] + pl.lpSum(mid_green_vars[h]) + pl.lpSum(low_green_vars[h])
        )
        model += total_green <= new_energy_supply[h]

    # 构建目标函数
    original_cost = pl.lpSum(
        [green_high[h] * new_energy_price[h] for h in hours]
        + [trad_high[h] * tradition_price[h] for h in hours]
        + [g * new_energy_price[h] for h in hours for g in mid_green_vars[h]]
        + [t * tradition_price[h] for h in hours for t in mid_trad_vars[h]]
        + [g * new_energy_price[h] for h in hours for g in low_green_vars[h]]
        + [t * tradition_price[h] for h in hours for t in low_trad_vars[h]]
    )

    green_total = (
        pl.lpSum(green_high[h] for h in hours)
        + pl.lpSum(g for h in hours for g in mid_green_vars[h])
        + pl.lpSum(g for h in hours for g in low_green_vars[h])
    )

    cost = original_cost - beta * green_total
    model += cost

    # 求解模型
    model.solve(pl.PULP_CBC_CMD(timeLimit=60, gapRel=0.01, strong=1, cuts=True))
    # 收集各小时绿色能源和总用电量数据
    hours = range(24)
    green_usage = {}
    total_usage = {}

    for h in hours:
        # 绿色能源使用量
        gh = pl.value(green_high[h])
        mid_green = sum(pl.value(g) for g in mid_green_vars[h])
        low_green = sum(pl.value(g) for g in low_green_vars[h])
        green_total = gh + mid_green + low_green

        # 传统能源使用量
        trad_high_val = pl.value(trad_high[h])
        mid_trad = sum(pl.value(t) for t in mid_trad_vars[h])
        low_trad = sum(pl.value(t) for t in low_trad_vars[h])
        trad_total = trad_high_val + mid_trad + low_trad

        # 总用电量
        total = green_total + trad_total

        green_usage[h] = green_total
        total_usage[h] = total

    return pl.value(model.objective), green_usage, total_usage


if __name__ == "__main__":
    trad_price, ne_price, ne_supply = load_attachment1("附件1_测试3.xlsx")
    high_tasks, mid_subtasks, low_subtasks = load_attachment2("附件2_测试3.xlsx")
    total_cost, green_usage, total_usage = build_and_solve_model(
        trad_price, ne_price, ne_supply, high_tasks, mid_subtasks, low_subtasks
    )
    print(f"优化后的总成本为：{total_cost:.2f}元")
    # 计算每小时绿色能源使用率
    hours = list(range(24))
    usage_rates = [
        (green_usage[h] / total_usage[h] * 100) if total_usage[h] > 0 else 0
        for h in hours
    ]
    trad_usage = {h: total_usage[h] - green_usage[h] for h in hours}
    # 生成折线图

    plt.figure(figsize=(12, 6))
    plt.plot(hours, usage_rates, marker="o", linestyle="-", color="#2ca02c")
    plt.title("每小时绿色能源使用率", fontsize=14)
    plt.xlabel("时间（时）", fontsize=12)
    plt.ylabel("绿色能源使用率（%）", fontsize=12)
    plt.xticks(hours)
    plt.ylim(0, 100)  # 固定y轴范围
    plt.grid(True, linestyle="--", alpha=0.7)
    plt.tight_layout()
    plt.savefig("green_energy_usage_rate.png")  # 保存图片
    plt.show()
    plt.figure(figsize=(12, 6))
    plt.bar(
        hours,
        trad_usage.values(),
        color="#ff7f0e",
        edgecolor="black",
        alpha=0.8,
        label="传统能源",
    )
    plt.title("每小时传统能源使用量", fontsize=14)
    plt.xlabel("时间（时）", fontsize=12)
    plt.ylabel("能源消耗量（千瓦时）", fontsize=12)
    plt.xticks(hours)
    plt.grid(axis="y", linestyle="--", alpha=0.7)
    plt.legend()
    plt.tight_layout()
    plt.savefig("traditional_energy_consumption.png")  # 保存图片
    plt.show()
    # 计算平均利用率
    average_usage = sum(usage_rates) / len(usage_rates)
    print(f"绿色能源平均利用率：{average_usage:.2f}%")
