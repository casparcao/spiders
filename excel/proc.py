import pandas as pd
import re

# 读取Excel
df = pd.read_excel("perform2.xlsx")


# 去除路径中的数字或UUID
def normalize_path1(path):
    return re.sub(r"[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}", "{id}", path)


def normalize_path2(path):
    return re.sub(r"[a-f0-9]{32}", "{id}", path)


def normalize_path3(path):
    return re.sub(r"[0-9]{2,}", "{id}", path)


def normalize_path4(path):
    return re.sub(r"[A-Z]{2,}", "{ref}", path)


df["normalized_path"] = df["请求路径"].apply(normalize_path1)\
    .apply(normalize_path2)\
    .apply(normalize_path3)\
    .apply(normalize_path4)

# 合并数据
result = df.groupby(["请求方式", "normalized_path"]).agg({
    "次数": "sum",
    "平均值": "mean",
    "50分位": "mean",
    "95分位": "mean"
}).reset_index()

# 输出结果
result.to_excel("merged_api_data.xlsx", index=False)