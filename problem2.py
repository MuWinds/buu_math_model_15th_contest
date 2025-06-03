import pandas as pd
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

    df_new_energy_price = pd.read_excel(file_path, sheet_name="新能源电价")
    new_energy_price = df_new_energy_price.set_index("时间（时）")[
        "电价（单位：元/千瓦时 ）"
    ].to_dict()

    df_supply = pd.read_excel(file_path, sheet_name="新能源电力供应量")
    new_energy_supply = (
        df_supply.set_index("时间（时）")["新能源电力供应（兆瓦）"] * 1000
    ).to_dict()

    return tradition_price, new_energy_price, new_energy_supply


def load_attachment2(file_path):
    """加载附件2数据并生成任务分配结构"""
    df = pd.read_excel(file_path, sheet_name="Sheet1")
    high_tasks = defaultdict(float)  # {小时: 任务量}
    mid_tasks = []  # (任务量, 发布时间)
    low_tasks = []

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

        # 高任务直接分配到各小时
        per_high = high / num_hours
        for h in hours:
            high_tasks[h] += per_high

        # 中低任务记录发布时间
        per_mid = mid / num_hours
        per_low = low / num_hours
        for h in hours:
            mid_tasks.append((per_mid, h))
            low_tasks.append((per_low, h))

    return high_tasks, mid_tasks, low_tasks


def calculate_cost(
    tradition_price,
    new_energy_price,
    new_energy_supply,
    high_tasks,
    mid_tasks,
    low_tasks,
):
    """新型调度策略的成本计算"""
    # 初始化数据结构
    remaining_green = defaultdict(float)
    cost = 0.0

    # 第一阶段：处理高优先级任务
    high_consumption = defaultdict(float)
    for hour in range(24):
        # 计算高任务电力需求
        energy_needed = high_tasks.get(hour, 0) * 80
        green_available = new_energy_supply.get(hour, 0)

        # 优先使用新能源
        green_used = min(energy_needed, green_available)
        trad_used = max(0, energy_needed - green_available)

        # 记录消耗和剩余
        high_consumption[hour] = (green_used, trad_used)
        remaining_green[hour] = green_available - green_used
        cost += green_used * new_energy_price[hour]
        cost += trad_used * tradition_price[hour]

    # 第二阶段：处理中优先级任务
    mid_green_usage = defaultdict(float)
    mid_trad_usage = defaultdict(float)
    for task in mid_tasks:
        task_energy = task[0] * 50
        publish_hour = task[1]

        # 生成允许的时间窗口
        allowed_hours = [(publish_hour + i) % 24 for i in range(24)]

        # 按绿色电价排序，优先低电价且剩余多的时段
        sorted_hours = sorted(
            allowed_hours, key=lambda h: (new_energy_price[h], -remaining_green[h])
        )

        # 尝试分配新能源
        allocated = False
        for h in sorted_hours:
            if remaining_green[h] >= task_energy:
                mid_green_usage[h] += task_energy
                remaining_green[h] -= task_energy
                allocated = True
                break

        # 新能源不足则找传统电价最低时段
        if not allocated:
            sorted_trad = sorted(allowed_hours, key=lambda h: tradition_price[h])
            for h in sorted_trad:
                mid_trad_usage[h] += task_energy
                allocated = True
                break

    # 第三阶段：处理低优先级任务
    low_green_usage = defaultdict(float)
    low_trad_usage = defaultdict(float)
    for task in low_tasks:
        task_energy = task[0] * 30
        publish_hour = task[1]

        # 生成允许的时间窗口
        allowed_hours = [(publish_hour + i) % 24 for i in range(24)]

        # 优先剩余新能源
        allocated = False
        for h in sorted(
            allowed_hours, key=lambda h: (new_energy_price[h], -remaining_green[h])
        ):
            if remaining_green[h] >= task_energy:
                low_green_usage[h] += task_energy
                remaining_green[h] -= task_energy
                allocated = True
                break

        # 否则找传统电价最低时段（凌晨低谷）
        if not allocated:
            sorted_trad = sorted(allowed_hours, key=lambda h: tradition_price[h])
            for h in sorted_trad:
                low_trad_usage[h] += task_energy
                allocated = True
                break

    # 计算总成本
    for hour in range(24):
        # 新能源部分
        green_used = mid_green_usage[hour] + low_green_usage[hour]
        cost += green_used * new_energy_price[hour]

        # 传统能源部分
        trad_used = mid_trad_usage[hour] + low_trad_usage[hour]
        cost += trad_used * tradition_price[hour]

    # 计算绿色能源使用率
    green_usage = {}
    for hour in range(24):
        # 获取各阶段新能源使用量
        high_green, high_trad = high_consumption.get(hour, (0, 0))
        mid_green = mid_green_usage.get(hour, 0)
        low_green = low_green_usage.get(hour, 0)

        # 获取各阶段传统能源使用量
        mid_trad = mid_trad_usage.get(hour, 0)
        low_trad = low_trad_usage.get(hour, 0)

        # 计算总量
        total_green = high_green + mid_green + low_green
        total_trad = high_trad + mid_trad + low_trad
        total_energy = total_green + total_trad

        # 计算使用率
        usage_rate = (total_green / total_energy * 100) if total_energy != 0 else 0.0
        green_usage[hour] = usage_rate

    # 计算传统能源使用量
    trad_usage = {}
    for hour in range(24):
        # 获取各阶段传统能源使用量
        high_trad = high_consumption.get(hour, (0, 0))[1]
        mid_trad = mid_trad_usage.get(hour, 0)
        low_trad = low_trad_usage.get(hour, 0)
        trad_usage[hour] = high_trad + mid_trad + low_trad  # 单位：千瓦时

    return cost, green_usage, trad_usage  # 返回新增的传统能源使用量


if __name__ == "__main__":
    # 加载数据
    trad_price, ne_price, ne_supply = load_attachment1("附件1_测试5.xlsx")
    high_tasks, mid_tasks, low_tasks = load_attachment2("附件2_测试5.xlsx")

    # 计算成本和绿色能源使用率
    total_cost, green_usage, trad_usage = calculate_cost(
        trad_price, ne_price, ne_supply, high_tasks, mid_tasks, low_tasks
    )

    # 绘制折线图
    hours = list(range(24))
    usage_rates = [green_usage[h] for h in hours]

    plt.figure(figsize=(12, 6))
    plt.plot(hours, usage_rates, marker="o", linestyle="-", color="#2ca02c")
    plt.title("每小时绿色能源使用率", fontsize=14, fontweight="bold")
    plt.xlabel("时间（时）", fontsize=12)
    plt.ylabel("绿色能源占比 (%)", fontsize=12)
    plt.xticks(hours)
    plt.grid(True, linestyle="--", alpha=0.7)
    plt.tight_layout()
    # 显示图表
    plt.show()
    # 计算平均利用率
    average_usage = sum(usage_rates) / len(usage_rates)
    print(f"绿色能源平均利用率：{average_usage:.2f}%")
    plt.figure(figsize=(12, 6))
    plt.bar(
        hours,
        [trad_usage[h] for h in hours],
        color="#1f77b4",
        edgecolor="black",
        alpha=0.7,
        label="传统能源使用量",
    )
    plt.title("每小时传统能源使用量（柱状图）", fontsize=14, fontweight="bold")
    plt.xlabel("时间（时）", fontsize=12)
    plt.ylabel("传统能源使用量（千瓦时）", fontsize=12)
    plt.xticks(hours)
    plt.grid(True, axis="y", linestyle="--", alpha=0.7)
    plt.tight_layout()
    plt.show()
    # 计算平均利用率
    average_usage = sum(usage_rates) / len(usage_rates)
    print(f"绿色能源平均利用率：{average_usage:.2f}%")
    # 打印总成本（保持原有输出）
    print(f"\n优化后的24小时总电力成本为：{total_cost:.2f}元")
