import pandas as pd
import re


def normalize_path_params(path):
    """
    将路径中的参数占位符统一替换为{id}格式
    
    参数:
    path: 原始路径字符串
    
    返回:
    格式化后的路径字符串
    """
    if not isinstance(path, str):
        return path
    
    # 将<xxx>形式替换为{id}
    path = re.sub(r'<([^>]+)>', r'{id}', path)
    
    # 将**形式替换为{id}
    path = re.sub(r'\*\*', r'{id}', path)
    
    # 将{xxxx}形式统一替换为{id}
    path = re.sub(r'\{[^}]+\}', r'{id}', path)
    
    return path


def find_common_apis(file1_path, file2_path, output_path):
    """
    找出两个Excel文件中都存在的API接口（交集）
    
    参数:
    file1_path: 第一个Excel文件路径
    file2_path: 第二个Excel文件路径
    output_path: 输出交集结果的Excel文件路径
    """
    print("正在读取第一个文件...")
    try:
        df1 = pd.read_excel(file1_path)
        # 将所有列内容转换为小写
        df1 = df1.apply(lambda x: x.astype(str).str.lower() if x.dtype == "object" else x)
        print(f"从 {file1_path} 读取到 {len(df1)} 条记录")
        print(f"列名: {list(df1.columns)}")
    except FileNotFoundError:
        print(f"错误：找不到文件 {file1_path}")
        return
    except Exception as e:
        print(f"读取文件 {file1_path} 时发生错误：{e}")
        return

    print("正在读取第二个文件...")
    try:
        df2 = pd.read_excel(file2_path)
        # 将所有列内容转换为小写
        df2 = df2.apply(lambda x: x.astype(str).str.lower() if x.dtype == "object" else x)
        print(f"从 {file2_path} 读取到 {len(df2)} 条记录")
        print(f"列名: {list(df2.columns)}")
    except FileNotFoundError:
        print(f"错误：找不到文件 {file2_path}")
        return
    except Exception as e:
        print(f"读取文件 {file2_path} 时发生错误：{e}")
        return

    # 确定用于比较的列
    # 假设至少有一个"请求路径"列
    path_column1 = None
    method_column1 = None
    path_column2 = None
    method_column2 = None

    # 查找第一个文件中的相关列
    for col in df1.columns:
        if '请求路径' in col or 'path' in col.lower():
            path_column1 = col
        elif '请求方法' in col or 'method' in col.lower():
            method_column1 = col

    # 查找第二个文件中的相关列
    for col in df2.columns:
        if '请求路径' in col or 'path' in col.lower():
            path_column2 = col
        elif '请求方法' in col or 'method' in col.lower():
            method_column2 = col

    print(f"文件1使用的列: 路径={path_column1}, 方法={method_column1}")
    print(f"文件2使用的列: 路径={path_column2}, 方法={method_column2}")

    # 标准化路径参数格式
    if path_column1:
        df1[path_column1] = df1[path_column1].apply(normalize_path_params)
        print("已完成文件1中的路径参数格式标准化")

    if path_column2:
        df2[path_column2] = df2[path_column2].apply(normalize_path_params)
        print("已完成文件2中的路径参数格式标准化")

    # 构建API标识（方法+路径）
    if method_column1 and path_column1:
        df1['API标识'] = df1[method_column1] + ':' + df1[path_column1]
    elif path_column1:
        df1['API标识'] = df1[path_column1]
    else:
        print("无法在文件1中找到合适的路径列")
        return

    if method_column2 and path_column2:
        df2['API标识'] = df2[method_column2] + ':' + df2[path_column2]
    elif path_column2:
        df2['API标识'] = df2[path_column2]
    else:
        print("无法在文件2中找到合适的路径列")
        return

    # 找出交集
    common_apis = df1[df1['API标识'].isin(df2['API标识'])]
    print(f"找到 {len(common_apis)} 个共同的API接口")

    # 显示前10个匹配的接口示例
    print("\n--- 共同的API接口示例 ---")
    for i, (_, row) in enumerate(common_apis.head(10).iterrows()):
        print(f"{i+1}. {row['API标识']}")
    print("------------------------")

    # 删除辅助列
    if 'API标识' in common_apis.columns:
        common_apis = common_apis.drop(columns=['API标识'])

    # 保存结果到Excel文件
    try:
        common_apis.to_excel(output_path, index=False)
        print(f"\n结果已保存到: {output_path}")
    except Exception as e:
        print(f"保存文件时发生错误：{e}")


if __name__ == "__main__":
    # 配置文件路径
    file1 = r"D:\codes\demo\spiders\testless\前端JS解析的接口汇总.xlsx"
    file2 = r"D:\codes\demo\spiders\testless\Swagger中定义的所有接口.xlsx"
    output_file = r"/testless/前端js解析的接口跟swagger定义接口的交集.xlsx"

    # 运行查找交集函数
    find_common_apis(file1, file2, output_file)