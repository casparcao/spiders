import pandas as pd
import os
import re
from openpyxl.styles import Font


def normalize_path_params(path):
    """
    将路径中的<xxx>或**等形式的参数占位符统一替换为{xxx}格式
    
    参数:
    path: 原始路径字符串
    
    返回:
    格式化后的路径字符串
    """
    if not isinstance(path, str):
        return path
    
    # 将<xxx>形式替换为{xxx}
    path = re.sub(r'<([^>]+)>', r'{\1}', path)
    
    # 将**形式替换为{id}（这里假设**代表一个ID参数）
    path = re.sub(r'\*\*', r'{id}', path)
    
    return path


def analyze_untested_apis(api_info_file, test_files, output_file):
    """
    分析未测试的API接口
    
    参数:
    api_info_file: 包含所有API接口定义的Excel文件路径
    test_files: 已测试接口的Excel文件列表
    output_file: 输出未测试接口到Excel文件的路径
    """
    
    # 读取所有API接口定义
    print("正在读取所有API接口定义...")
    try:
        api_info = pd.read_excel(api_info_file)
        print(f"总共读取到 {len(api_info)} 个API接口")
        print(f"API信息列名: {list(api_info.columns)}")
    except FileNotFoundError:
        print(f"错误：找不到文件 {api_info_file}")
        return
    except Exception as e:
        print(f"读取API信息时发生错误：{e}")
        return
    
    # 标准化API定义文件中的路径参数格式
    if '请求路径' in api_info.columns:
        api_info['原始请求路径'] = api_info['请求路径']  # 保存原始路径
        api_info['请求路径'] = api_info['请求路径'].apply(normalize_path_params)
        print("已完成API定义文件中的路径参数格式标准化")
    
    # 读取已测试的接口
    tested_apis = pd.DataFrame()
    for test_file in test_files:
        try:
            df = pd.read_excel(test_file)
            tested_apis = pd.concat([tested_apis, df], ignore_index=True)
            print(f"从 {test_file} 读取到 {len(df)} 个已测试接口")
        except FileNotFoundError:
            print(f"警告：找不到测试文件 {test_file}")
        except Exception as e:
            print(f"读取测试文件 {test_file} 时发生错误：{e}")
    
    print(f"总共读取到 {len(tested_apis)} 个已测试接口")
    
    # 标准化测试文件中的路径参数格式
    if '请求路径' in tested_apis.columns:
        tested_apis['原始请求路径'] = tested_apis['请求路径']  # 保存原始路径
        tested_apis['请求路径'] = tested_apis['请求路径'].apply(normalize_path_params)
        print("已完成测试文件中的路径参数格式标准化")
    
    # 确定用于比较的列
    api_path_column = None
    api_method_column = None
    api_service_column = None
    
    # API定义文件中的列名
    if '请求路径' in api_info.columns:
        api_path_column = '请求路径'
    if '请求方法' in api_info.columns:
        api_method_column = '请求方法'
    if '服务名称' in api_info.columns:
        api_service_column = '服务名称'
    
    test_path_column = None
    test_method_column = None
    test_service_in_path = False  # 路径中是否包含服务名称
    
    # 测试文件中的列名（假设有统一格式）
    if '请求路径' in tested_apis.columns:
        test_path_column = '请求路径'
    if '请求方法' in tested_apis.columns:
        test_method_column = '请求方法'
    if '请求方式' in tested_apis.columns:
        test_method_column = '请求方式'

    print(f"API定义文件使用列: 路径={api_path_column}, 方法={api_method_column}, 服务={api_service_column}")
    print(f"测试文件使用列: 路径={test_path_column}, 方法={test_method_column}")
    
    # 构建完整的API标识（方法+路径）
    # 对于API定义文件，需要组合服务名称和路径
    if api_service_column and api_path_column and api_method_column:
        api_info['完整路径'] = '/' + api_info[api_service_column] + api_info[api_path_column]
        api_info['API标识'] = api_info[api_method_column] + ':' + api_info['完整路径']
        print("构建API定义文件的完整标识完成")
    
    # 对于测试文件，路径应该已经包含服务名称前缀
    if test_path_column:
        if test_method_column:
            tested_apis['API标识'] = tested_apis[test_method_column] + ':' + tested_apis[test_path_column]
        else:
            # 如果测试文件中没有方法列，只使用路径
            tested_apis['API标识'] = tested_apis[test_path_column]
        print("构建测试文件的API标识完成")
    
    # 去重已测试的接口
    tested_apis_unique = tested_apis.drop_duplicates(subset=['API标识'])
    print(f"去重后共有 {len(tested_apis_unique)} 个唯一已测试接口")
    
    # 找出未测试的接口
    print("正在查找未测试的接口...")
    untested_apis = api_info[~api_info['API标识'].isin(tested_apis_unique['API标识'])]
    print(f"发现 {len(untested_apis)} 个未测试的接口")
    
    # 选择需要输出的列（去除我们添加的辅助列）
    output_columns = [col for col in api_info.columns if col not in ['完整路径', 'API标识']]
    untested_apis_output = untested_apis[output_columns]
    
    # 创建Excel写入器
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 按服务分组并将每个服务写入单独的sheet
        if api_service_column:
            for service_name, group in untested_apis_output.groupby(api_service_column):
                # 将数据写入对应服务名称的sheet
                sheet_name = str(service_name)[:31]  # Excel sheet名称限制为31个字符
                group.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # 设置字体为微软雅黑
                worksheet = writer.sheets[sheet_name]
                for row in worksheet.iter_rows():
                    for cell in row:
                        cell.font = Font(name='微软雅黑')
                print(f"服务 '{service_name}' 的未测试接口已写入sheet '{sheet_name}'")
        
        # 写入汇总sheet
        untested_apis_output.to_excel(writer, sheet_name='全部未测试接口', index=False)
        
        # 设置字体为微软雅黑
        worksheet = writer.sheets['全部未测试接口']
        for row in worksheet.iter_rows():
            for cell in row:
                cell.font = Font(name='微软雅黑')
        print("所有未测试接口已写入sheet '全部未测试接口'")
    
    print(f"结果已保存到: {output_file}")


if __name__ == "__main__":
    # 配置文件路径
    api_info_file = "api_info.xlsx"  # 所有API接口定义文件
    test_files = [  # 测试过的接口文件
        "first.xlsx",
        "second.xlsx", 
        "third.xlsx"
    ]
    output_file = "untested_apis.xlsx"  # 输出文件
    
    # 运行分析
    analyze_untested_apis(api_info_file, test_files, output_file)