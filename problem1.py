import pandas as pd
import matplotlib.pyplot as plt

# 设置中文显示
plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False


# 读取附件1数据
def load_attachment1(file_path):
    # 读取传统电价
    df_tradition = pd.read_excel(file_path, sheet_name="传统电价")
    tradition_price = df_tradition.set_index("时间（时）")[
        "传统电价（单位：元/千瓦时 ）"
    ].to_dict()
    print(tradition_price)

    # 读取新能源电价
    df_new_energy_price = pd.read_excel(file_path, sheet_name="新能源电价")
    new_energy_price = df_new_energy_price.set_index("时间（时）")[
        "电价（单位：元/千瓦时 ）"
    ].to_dict()

    # 读取新能源供应量并转换单位（兆瓦->千瓦时）
    df_supply = pd.read_excel(file_path, sheet_name="新能源电力供应量")
    new_energy_supply = (
        df_supply.set_index("时间（时）")["新能源电力供应（兆瓦）"] * 1000
    ).to_dict()

    return tradition_price, new_energy_price, new_energy_supply


# 读取附件2数据并生成小时任务分配
def load_attachment2(file_path):
    df = pd.read_excel(file_path, sheet_name="Sheet1")
    hours_tasks = {hour: {"high": 0.0, "mid": 0.0, "low": 0.0} for hour in range(24)}

    for _, row in df.iterrows():
        time_range = row.iloc[0]
        tasks = {"high": row.iloc[1], "mid": row.iloc[2], "low": row.iloc[3]}

        # 解析时间段
        start_str, end_str = time_range.split("-")
        start_hour = int(start_str.split(":")[0])
        end_hour = int(end_str.split(":")[0])
        hours = list(range(start_hour, end_hour))
        num_hours = len(hours)

        if num_hours == 0:
            continue

        # 分配任务到每个小时
        for task_type in ["high", "mid", "low"]:
            per_hour = tasks[task_type] / num_hours
            for h in hours:
                hours_tasks[h][task_type] += per_hour

    return hours_tasks


def calculate_cost(
    tradition_price: dict,
    new_energy_price: dict,
    new_energy_supply: dict,
    hours_tasks: dict,
):
    total_cost = 0.0
    hourly_usage_rates = []  # 新增：存储每小时的绿色能源使用率
    traditional_usage = []  # 新增：存储每小时传统能源使用量（千瓦时）

    for hour in range(24):
        # 获取当前小时参数
        tp_price = tradition_price.get(hour, 0)
        np_price = new_energy_price.get(hour, 0)
        supply = new_energy_supply.get(hour, 0)

        # 计算总能耗
        tasks = hours_tasks[hour]
        energy = tasks["high"] * 80 + tasks["mid"] * 50 + tasks["low"] * 30

        # 计算绿色能源使用率
        if energy == 0:
            usage_rate = 0.0
        else:
            usage_rate = min(supply / energy, 1.0)
        hourly_usage_rates.append(usage_rate * 100)  # 转换为百分比

        # 计算成本
        if energy <= supply:
            cost = energy * np_price
        else:
            cost = supply * np_price + (energy - supply) * tp_price
        # 计算传统能源用量
        if energy > supply:
            trad_energy = energy - supply
        else:
            trad_energy = 0
        traditional_usage.append(trad_energy)
        total_cost += cost

    return total_cost, hourly_usage_rates, traditional_usage  # 修改返回值


# 主程序
if __name__ == "__main__":
    # 加载数据
    trad_price, ne_price, ne_supply = load_attachment1("附件1_测试5.xlsx")
    task_distribution = load_attachment2("附件2_测试5.xlsx")

    total, usage_rates, traditional_usage = calculate_cost(
        trad_price, ne_price, ne_supply, task_distribution
    )

    plt.figure(figsize=(12, 6))
    plt.bar(range(24), traditional_usage, color="#1f77b4", alpha=0.7)
    plt.title("每小时传统能源使用量", fontsize=14)
    plt.xlabel("时间（时）", fontsize=12)
    plt.ylabel("用电量 (千瓦时)", fontsize=12)
    plt.xticks(range(24))
    plt.grid(True, axis="y", alpha=0.3)

    # 生成折线图
    plt.figure(figsize=(12, 6))
    plt.plot(range(24), usage_rates, marker="o", linestyle="-", color="#2ca02c")
    plt.title("每小时绿色能源使用率", fontsize=14)
    plt.xlabel("时间（时）", fontsize=12)
    plt.ylabel("使用率 (%)", fontsize=12)
    plt.xticks(range(24))
    plt.ylim(0, 105)
    plt.grid(True, alpha=0.3)

    # 标记最高/最低点
    max_rate = max(usage_rates)
    min_rate = min(usage_rates)
    plt.scatter(usage_rates.index(max_rate), max_rate, color="red", zorder=5)
    plt.scatter(usage_rates.index(min_rate), min_rate, color="blue", zorder=5)

    plt.tight_layout()
    plt.show()
    # 计算平均利用率
    average_usage = sum(usage_rates) / len(usage_rates)
    print(f"绿色能源平均利用率：{average_usage:.2f}%")
    # 输出结果
    print(f"24小时总电力成本为：{total:.2f} 元")
