import pandas as pd


def merge_excel_files(file1_path, file2_path, output_path):
    """
    合并两个Excel文件并去重
    """
    # 读取两个Excel文件
    df1 = pd.read_excel(file1_path)
    df2 = pd.read_excel(file2_path)
    
    # 将所有字符串列转换为小写
    for col in df1.columns:
        if df1[col].dtype == 'object':  # 只对字符串类型列进行转换
            df1[col] = df1[col].str.lower()
    
    for col in df2.columns:
        if df2[col].dtype == 'object':  # 只对字符串类型列进行转换
            df2[col] = df2[col].str.lower()
    
    # 合并两个DataFrame
    merged_df = pd.concat([df1, df2], ignore_index=True)
    
    # 基于所有列去重
    merged_df = merged_df.drop_duplicates().reset_index(drop=True)
    
    # 保存到新的Excel文件
    merged_df.to_excel(output_path, index=False, sheet_name='Merged APIs')
    
    print(f"✅ 合并完成，原始文件1包含 {len(df1)} 条记录，文件2包含 {len(df2)} 条记录")
    print(f"合并后共有 {len(merged_df)} 条唯一记录，结果保存到 '{output_path}'")
    
    # 显示前10条记录预览
    print("\n合并后前10条记录预览：")
    print(merged_df.head(10))
    
    return merged_df


def main():
    file1 = 'normalized_apis3_fixed.xlsx'
    file2 = 'normalized_apis2_fixed.xlsx'
    output_file = 'merged_normalized_apis.xlsx'
    
    try:
        merge_excel_files(file1, file2, output_file)
    except FileNotFoundError as e:
        print(f"❌ 文件未找到: {e}")
    except Exception as e:
        print(f"❌ 处理过程中出现错误: {e}")


if __name__ == '__main__':
    main()