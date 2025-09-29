import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

if not os.path.exists("api_performance_charts"):
    os.makedirs("api_performance_charts")

# 设置中文字体和绘图风格
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False
sns.set_style("whitegrid")

# 示例：加载数据（请替换为你的文件路径）
df = pd.read_csv('api_logs.csv')  # 替换为实际路径

# 构建接口唯一标识
df['api_signature'] = df['url_template'] + " [" + df['request_method'] + "]"

# 按接口聚合基础统计
summary = df.groupby('api_signature').agg(
    count=('duration', 'count'),
    mean_duration=('duration', 'mean'),
    std_duration=('duration', 'std'),  # 用于稳定性分析
    min_duration=('duration', 'min'),
    max_duration=('duration', 'max'),
    p95_duration=('duration', lambda x: x.quantile(0.95))
).reset_index()

# 补充模块信息（可选）
module_map = df.groupby('api_signature')['module'].first()
summary['module'] = summary['api_signature'].map(module_map)

# 处理标准差为 NaN 的情况（单次请求）
summary['std_duration'] = summary['std_duration'].fillna(0)


# 定义请求次数区间
# bins = [0, 100, 200, 500, 1000, float('inf')]
# labels = ['1-100', '101-200', '201-500', '501-1000', '1000+']
bins = [200, float('inf')]
labels = ['200+']
summary['count_interval'] = pd.cut(summary['count'], bins=bins, labels=labels, right=True)

# Step 5: 对每个「模块 + 请求次数区间」组合绘图
for (module, interval), group in summary.groupby(['module', 'count_interval']):
    if len(group) == 0:
        continue
    # 取该组合下平均响应时间最慢的 Top 20 接口
    top_slow = group.nlargest(20, 'mean_duration')
    # 绘图
    plt.figure(figsize=(10, 6))
    sns.barplot(data=top_slow, x='mean_duration', y='api_signature', palette='Reds_r', hue='api_signature', legend=False)
    plt.title(f"Top 20 Slowest APIs | Module: {module} | Request Count: {interval}")
    plt.xlabel("Average Response Time (ms)")
    plt.ylabel("API Endpoint")
    plt.tight_layout()
    # plt.savefig(f"api_performance_charts/slowest_{module}_{interval}.png", dpi=150, bbox_inches='tight')
    plt.close()
    # plt.show()


# =============================
# 自定义颜色映射函数
# =============================
def get_color(duration):
    """根据响应时间返回对应的颜色"""
    if duration < 100:
        # 绿色系：0ms -> 深绿，100ms -> 浅绿
        ratio = duration / 100
        r = 0.8 - 0.5 * ratio
        g = 0.9 - 0.3 * ratio
        b = 0.8 - 0.7 * ratio
        return (r, g, b)  # 深绿到浅绿
    elif duration < 200:
        # 蓝色系：100ms -> 浅蓝，200ms -> 深蓝
        ratio = (duration - 100) / 100
        r = 0.8 - 0.6 * ratio
        g = 0.9 - 0.8 * ratio
        b = 1.0 - 0.2 * ratio
        return (r, g, b)
    elif duration < 500:
        # 橘色系：200ms -> 浅橘，500ms -> 深橘
        ratio = (duration - 200) / 300
        r = 1.0
        g = 0.7 - 0.5 * ratio
        b = 0.3 - 0.2 * ratio
        return (r, g, b)
    else:
        # 红色系：500ms+ -> 深红
        ratio = (duration - 500) / 500 if duration < 1000 else 1  # 限制最大值
        r = 1.0
        g = 0.3 - 0.2 * ratio
        b = 0.3 - 0.2 * ratio
        return (r, g, b)


# =============================
# 对每个「模块 + 请求次数区间」组合绘图
# =============================
for (module, interval), group in summary.groupby(['module', 'count_interval']):
    if len(group) == 0:
        continue

    # 取该组合下平均响应时间最慢的 Top 20 接口
    top_slow = group.nlargest(20, 'mean_duration').copy()

    # 为每个接口计算颜色
    top_slow['color'] = top_slow['mean_duration'].apply(get_color)

    # 排序：保持从慢到快（倒序），但颜色独立
    top_slow = top_slow.sort_values('mean_duration', ascending=False)

    # 绘图
    plt.figure(figsize=(12, 8))
    bars = plt.barh(
        y=top_slow['api_signature'],
        width=top_slow['mean_duration'],
        color=top_slow['color'].tolist(),
        edgecolor='silver',
        linewidth=0.5
    )

    plt.title(f"Top 20 Slowest APIs | Module: {module} | Request Count: {interval}")
    plt.xlabel("Average Response Time (ms)")
    plt.ylabel("API Endpoint")

    # 反转 Y 轴，使最慢的在最上面
    plt.gca().invert_yaxis()

    # 添加数值标签（可选）
    for i, (idx, row) in enumerate(top_slow.iterrows()):
        plt.text(row['mean_duration'] + 10, i, f"{row['mean_duration']:.0f}ms",
                 va='center', fontsize=9, color='black', alpha=0.8)

    plt.tight_layout()
    safe_module = str(module).replace("/", "_").replace("\\", "_").replace(" ", "_")
    plt.savefig(f"api_performance_charts/slowest_{safe_module}_{interval}.png", dpi=150, bbox_inches='tight')
    plt.close()

    print(f"✅ 已生成模块 [{module}] 的 Top 20 最慢接口图（彩色编码）")


# 方法1：直接用标准差
summary['duration_std'] = summary['std_duration']

# 方法2（推荐）：用变异系数（消除量纲影响）
summary['cv'] = summary['std_duration'] / (summary['mean_duration'] + 1e-6)  # 防止除零

summary_filtered = summary[summary['count'] > 100].copy()

if len(summary_filtered) == 0:
    print("⚠️  没有找到调用次数大于 100 的接口，无法分析不稳定性。")
else:
    # =============================
    # 按 module 分组：找出每个模块中 CV 最高的 Top 20 接口（仅限 count > 100）
    # =============================
    for module_name, module_group in summary_filtered.groupby('module'):
        if len(module_group) == 0:
            continue

        # 取该模块下 CV 最高的前 20 个接口
        top_unstable_apis = module_group.nlargest(20, 'cv')['api_signature']

        if len(top_unstable_apis) == 0:
            print(f"⚠️  模块 [{module_name}] 中没有满足 count > 100 且 CV 有效的接口。")
            continue

        # 从原始数据中筛选这些接口的所有请求记录
        filtered_df = df[df['api_signature'].isin(top_unstable_apis)]

        if len(filtered_df) == 0:
            continue

        # 绘制箱线图
        plt.figure(figsize=(12, 8))
        sns.boxplot(
            data=filtered_df,
            x='duration',
            y='api_signature',
            palette='Pastel1',
            hue='api_signature',
            legend=False,
            showfliers=False  # 可选：隐藏异常值点，提升可读性
        )
        plt.title(f"Top 20 Most Unstable APIs (CV) | Module: {module_name} | Only APIs with >100 Calls")
        plt.xlabel("Response Time (ms)")
        plt.ylabel("API Endpoint")
        plt.tight_layout()

        # 保存图像
        safe_module_name = str(module_name).replace("/", "_").replace("\\", "_")
        plt.savefig(f"api_performance_charts/unstable_boxplot_{safe_module_name}.png", dpi=150, bbox_inches='tight')
        plt.close()

        print(f"✅ 已生成模块 [{module_name}] 的不稳定接口箱线图（仅含调用 >100 次的接口）")