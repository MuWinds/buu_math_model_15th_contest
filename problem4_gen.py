import pandas as pd
import numpy as np
import os

os.makedirs("鲁棒性测试数据", exist_ok=True)


def process_attachment1():
    # 完整示例数据结构
    trad_price = {
        "时间（时）": list(range(24)),
        "传统电价（单位：元/千瓦时 ）": [
            0.5,
            0.5,
            0.5,
            0.5,
            0.6,
            0.7,
            0.8,
            0.9,
            1.0,
            1.2,
            1.3,
            1.3,
            1.2,
            1.1,
            1.0,
            1.0,
            1.1,
            1.2,
            1.3,
            1.2,
            1.1,
            1.0,
            0.8,
            0.6,
        ],
    }

    new_energy_price = {
        "时间（时）": list(range(24)),
        "电价（单位：元/千瓦时 ）": [
            0.6,
            0.6,
            0.6,
            0.6,
            0.5,
            0.5,
            0.4,
            0.4,
            0.4,
            0.3,
            0.3,
            0.3,
            0.3,
            0.4,
            0.5,
            0.5,
            0.5,
            0.5,
            0.6,
            0.6,
            0.6,
            0.6,
            0.6,
            0.6,
        ],
    }

    new_energy_supply = {
        "时间（时）": list(range(24)),
        "新能源电力供应（兆瓦）": [
            0,
            0,
            0,
            0,
            0.5,
            1.4,
            1.8,
            2.1,
            2.4,
            2.4,
            2.8,
            3.2,
            3.4,
            3.3,
            3.1,
            2.9,
            2.6,
            2.5,
            2.3,
            1.5,
            1.0,
            0,
            0,
            0,
        ],
    }

    for i in range(1, 11):
        # 创建所有副本
        df_trad = pd.DataFrame(trad_price)
        df_price = pd.DataFrame(new_energy_price)
        df_supply = pd.DataFrame(new_energy_supply)

        # 传统电价全局波动
        trad_fluctuation = np.random.uniform(0.8, 1.2)
        df_trad["传统电价（单位：元/千瓦时 ）"] = (
            df_trad["传统电价（单位：元/千瓦时 ）"] * trad_fluctuation
        )
        df_trad["传统电价（单位：元/千瓦时 ）"] = df_trad[
            "传统电价（单位：元/千瓦时 ）"
        ].round(2)

        # 新能源供应量逐小时波动
        for idx in df_supply.index:
            val = df_supply.at[idx, "新能源电力供应（兆瓦）"]
            if val == 0:
                continue
            # 随机选择波动幅度
            if np.random.rand() > 0.5:
                fluctuation = np.random.uniform(-0.1, 0.1)
            else:
                fluctuation = np.random.uniform(-0.2, 0.2)
            new_val = max(0, val * (1 + fluctuation))
            df_supply.at[idx, "新能源电力供应（兆瓦）"] = round(new_val, 1)

        # 保存文件（包含三个sheet）
        with pd.ExcelWriter(f"鲁棒性测试数据/附件1_测试{i}.xlsx") as writer:
            df_trad.to_excel(writer, sheet_name="传统电价", index=False)
            df_price.to_excel(writer, sheet_name="新能源电价", index=False)
            df_supply.to_excel(writer, sheet_name="新能源电力供应量", index=False)


def process_attachment2():
    tasks = {
        "时间": [
            "00:00-06:00",
            "06:00-08:00",
            "08:00-12:00",
            "12:00-14:00",
            "14:00-18:00",
            "18:00-22:00",
            "22:00-24:00",
        ],
        "高紧急任务数": [0, 0, 114, 54, 152, 50, 0],
        "中紧急任务数": [40, 55, 72, 95, 80, 50, 20],
        "低紧急任务数": [60, 70, 0, 0, 0, 40, 15],
    }

    for i in range(1, 11):
        df = pd.DataFrame(tasks)

        # 生成缩放因子
        scale = np.random.uniform(0.85, 1.15)

        # 应用缩放并取整
        for col in ["高紧急任务数", "中紧急任务数", "低紧急任务数"]:
            df[col] = (df[col] * scale).round().astype(int)
            df[col] = df[col].apply(lambda x: max(0, x))

        # 保存文件
        df.to_excel(f"鲁棒性测试数据/附件2_测试{i}.xlsx", index=False)


# 执行生成
process_attachment1()
process_attachment2()
print("数据生成完成，请查看'鲁棒性测试数据'目录")
