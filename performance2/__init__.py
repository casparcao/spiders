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


def plot_top_slowest(df, metric='mean_duration', title=""):
    title_key="plot_top_slowest"
    top = df.nlargest(20, metric)
    plt.figure(figsize=(10, 6))
    sns.barplot(data=top, x=metric, y='api_signature', palette='Reds_r', hue='api_signature', legend=False)
    plt.title(f"Top 20 {title}")
    plt.xlabel("Average Response Time (ms)")
    plt.ylabel("API Endpoint")
    plt.tight_layout()
    plt.savefig(f"api_performance_charts/{title_key}.png", dpi=150, bbox_inches='tight')
    plt.show()


plot_top_slowest(summary, 'mean_duration', "Slowest APIs by Average Response Time")


# 定义请求次数区间
bins = [0, 100, 200, 500, 1000, float('inf')]
labels = ['1-100', '101-200', '201-500', '501-1000', '1000+']
summary['count_interval'] = pd.cut(summary['count'], bins=bins, labels=labels, right=True)

# Step 5: 对每个「模块 + 请求次数区间」组合绘图
for (module, interval), group in summary.groupby(['module', 'count_interval']):
    if len(group) == 0:
        continue
    # 取该组合下平均响应时间最慢的 Top 20 接口
    top_slow = group.nlargest(20, 'mean_duration')
    # 绘图
    plt.figure(figsize=(10, 6))
    sns.barplot(data=top_slow, x='mean_duration', y='api_signature', palette='Reds_r')
    plt.title(f"Top 20 Slowest APIs | Module: {module} | Request Count: {interval}")
    plt.xlabel("Average Response Time (ms)")
    plt.ylabel("API Endpoint")
    plt.tight_layout()
    plt.show()


def plot_by_request_count_range(data, low, high, label):
    subset = data[(data['count'] > low) & (data['count'] <= high)]
    if len(subset) == 0:
        print(f"No APIs found in request count range {low}-{high}")
        return
    title_key=f"plot_top_slowest_{label}"
    top = subset.nlargest(20, 'mean_duration')
    plt.figure(figsize=(10, 6))
    sns.barplot(data=top, x='mean_duration', y='api_signature', palette='Oranges_r', hue='api_signature', legend=False)
    plt.title(f"Top 20 Slowest APIs | Request Count: {low+1} - {high}")
    plt.xlabel("Average Response Time (ms)")
    plt.ylabel("API Endpoint")
    plt.tight_layout()
    plt.savefig(f"api_performance_charts/{title_key}.png", dpi=150, bbox_inches='tight')
    plt.show()


# 调用绘制各区间
plot_by_request_count_range(summary, 0, 100, "0-100")
plot_by_request_count_range(summary, 100, 200, "101-200")
plot_by_request_count_range(summary, 200, 500, "201-500")
plot_by_request_count_range(summary, 500, 1000, "501-1000")
plot_by_request_count_range(summary, 1000, float('inf'), "1000+")


top_freq = summary.nlargest(20, 'count')
plt.figure(figsize=(10, 6))
sns.barplot(data=top_freq, x='count', y='api_signature', palette='Blues_r', hue='api_signature', legend=False)
plt.title("Top 20 Most Frequently Called APIs")
plt.xlabel("Request Count")
plt.ylabel("API Endpoint")
plt.tight_layout()
plt.savefig(f"api_performance_charts/top_freq.png", dpi=150, bbox_inches='tight')
plt.show()


top_fast = summary.nsmallest(20, 'mean_duration')
plt.figure(figsize=(10, 6))
sns.barplot(data=top_fast, x='mean_duration', y='api_signature', palette='Greens', hue='api_signature', legend=False)
plt.title("Top 20 Fastest APIs by Average Response Time")
plt.xlabel("Average Response Time (ms)")
plt.ylabel("API Endpoint")
plt.tight_layout()
plt.savefig(f"api_performance_charts/top_fast.png", dpi=150, bbox_inches='tight')
plt.show()


def plot_fastest_by_range(data, low, high, label):
    subset = data[(data['count'] > low) & (data['count'] <= high)]
    if len(subset) == 0:
        print(f"No APIs found in request count range {low+1}-{high}")
        return
    top = subset.nsmallest(20, 'mean_duration')
    plt.figure(figsize=(10, 6))
    sns.barplot(data=top, x='mean_duration', y='api_signature', palette='Greens', hue='api_signature', legend=False)
    plt.title(f"Top 20 Fastest APIs | Request Count: {low+1} - {high}")
    plt.xlabel("Average Response Time (ms)")
    plt.ylabel("API Endpoint")
    plt.tight_layout()
    plt.savefig(f"api_performance_charts/plot_fastest_{label}.png", dpi=150, bbox_inches='tight')
    plt.show()


# 调用
plot_fastest_by_range(summary, 0, 100, "0-100")
plot_fastest_by_range(summary, 100, 200, "101-200")
plot_fastest_by_range(summary, 200, 500, "201-500")
plot_fastest_by_range(summary, 500, 1000, "501-1000")
plot_fastest_by_range(summary, 1000, float('inf'), "1000+")


# 方法1：直接用标准差
summary['duration_std'] = summary['std_duration']

# 方法2（推荐）：用变异系数（消除量纲影响）
summary['cv'] = summary['std_duration'] / (summary['mean_duration'] + 1e-6)  # 防止除零

# 按变异系数排序（波动相对最大）
unstable = summary.nlargest(20, 'cv')

plt.figure(figsize=(10, 6))
sns.barplot(data=unstable, x='cv', y='api_signature', palette='Purples_r', hue='api_signature', legend=False)
plt.title("Top 20 APIs with Most Unstable Response Time (Highest CV)")
plt.xlabel("Coefficient of Variation (Std/Mean)")
plt.ylabel("API Endpoint")
plt.tight_layout()
plt.savefig(f"api_performance_charts/unstable_rt.png", dpi=150, bbox_inches='tight')
plt.show()

# 示例：按小时和星期几的热力图
df['hour'] = pd.to_datetime(df['request_time']).dt.hour
df['weekday'] = pd.to_datetime(df['request_time']).dt.weekday  # 0=周一, 6=周日

pivot = df.groupby(['weekday', 'hour'])['duration'].mean().unstack()
plt.figure(figsize=(12, 6))
sns.heatmap(pivot, annot=True, fmt=".0f", cmap="YlOrRd")
plt.title("Average Response Time by Hour and Weekday")
plt.ylabel("Day of Week")
plt.xlabel("Hour of Day")
plt.show()


# 添加错误率字段
summary['error_rate'] = df.groupby('api_signature')['response_code'].apply(lambda x: (x >= 500).mean()).values

# 双Y轴图示例
fig, ax1 = plt.subplots(figsize=(10, 6))
top20 = summary.nlargest(20, 'mean_duration')
ax1.bar(top20['api_signature'], top20['mean_duration'], color='skyblue', label='Avg Duration')
ax2 = ax1.twinx()
ax2.plot(top20['api_signature'], top20['error_rate'], color='red', marker='o', label='Error Rate')
ax1.set_ylabel("Avg Response Time (ms)")
ax2.set_ylabel("Error Rate")
ax1.set_title("Top 20 Slow APIs: Latency vs Error Rate")
ax1.tick_params(axis='x', rotation=45)
plt.show()

# 客户端维度聚合
client_api = df.groupby(['client_id', 'api_signature'])['duration'].agg(['mean', 'count']).reset_index()
client_summary = df.groupby('client_id').agg(
    total_calls=('duration', 'count'),
    avg_latency=('duration', 'mean'),
    error_rate=('response_code', lambda x: (x >= 500).mean()),
    max_single_delay=('duration', 'max')
).reset_index()

# 绘图：客户端调用总数 TOP20
top_clients = client_summary.nlargest(20, 'total_calls')
sns.barplot(data=top_clients, x='total_calls', y='client_id', palette='Blues_r', hue="client_id", legend=False)
plt.title("Top 20 Clients by Call Volume")
plt.show()


module_perf = df.groupby('module').agg(
    avg_duration=('duration', 'mean'),
    p95_duration=('duration', lambda x: x.quantile(0.95)),
    error_rate=('response_code', lambda x: (x >= 500).mean()),
    stability=('duration', lambda x: x.std() / (x.mean() + 1))  # CV
).reset_index()

# 简单评分（越低越好）
module_perf['score'] = (
    (module_perf['avg_duration'] / 100) +
    (module_perf['p95_duration'] / 200) +
    (module_perf['error_rate'] * 100) +
    (module_perf['stability'] * 50)
)

# 排名
module_perf = module_perf.sort_values('score')
sns.barplot(data=module_perf, x='score', y='module', palette='Reds_r', hue="module", legend=False)
plt.title("Module Performance Health Score")
plt.show()


plt.figure(figsize=(10, 6))
sns.scatterplot(data=summary, x='count', y='mean_duration', hue='module', alpha=0.7)
sns.regplot(data=summary, x='count', y='mean_duration', scatter=False, color='red')
plt.title("Call Frequency vs Average Latency")
plt.xscale('log')  # 对数坐标更清晰
plt.yscale('log')
plt.show()


p99 = df['duration'].quantile(0.99)
slow_calls = df[df['duration'] > p99]

# 分析慢请求的客户端分布
top_clients_in_slow = slow_calls['client_id'].value_counts().head(10)
top_clients_in_slow.plot(kind='bar', title="Top Clients in P99+ Slow Requests")
plt.show()


# 标记是否达标
summary['sla_compliant'] = summary['p95_duration'] < 500  # 假设 SLA 是 500ms
compliance_rate = summary['sla_compliant'].mean()

# 按模块统计合规率
module_sla = df.merge(summary[['api_signature', 'sla_compliant']], on='api_signature') \
               .groupby('module')['sla_compliant'].mean()

module_sla.plot(kind='bar', title="SLA Compliance Rate by Module (P95 < 500ms)")
plt.ylabel("Compliance Rate")
plt.show()


# 针对每个模块进行处理
for module, group in df.groupby('module'):
    # 计算每个 api 的请求数量并排序，选择前 20
    top_apis = group['api_signature'].value_counts().nlargest(20).index
    # 筛选出属于这些 api 的记录
    filtered_data = group[group['api_signature'].isin(top_apis)]
    # 设置绘图大小
    plt.figure(figsize=(12, 8))
    # 绘制箱线图
    sns.boxplot(x='api_signature', y='duration', data=filtered_data)
    # 添加标题和轴标签
    plt.title(f'Top 20 APIs by Request Count in Module {module}')
    plt.xlabel('API Signature')
    plt.ylabel('Response Time (ms)')
    # 自动调整 x 轴刻度标签的角度以提高可读性
    plt.xticks(rotation=45, ha='right')
    # 显示图形
    plt.tight_layout()
    plt.show()



import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

st.title("API Performance Analyzer")

df = pd.read_csv("summary.csv")
st.dataframe(df)

option = st.selectbox("Select Chart", ["Top Slow", "By Client", "Time Trend"])

if option == "Top Slow":
    top = df.nlargest(10, 'mean_duration')
    fig, ax = plt.subplots()
    sns.barplot(data=top, x='mean_duration', y='api_signature', ax=ax, hue='api_signature', legend=False)
    st.pyplot(fig)