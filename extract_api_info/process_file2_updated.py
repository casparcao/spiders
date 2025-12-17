import re
import pandas as pd


def normalize_path(raw_path):
    """
    标准化 API 路径：
    1. 移除 ? 及之后的查询参数
    2. 将 ${...} 替换为 {id}
    3. 移除路径末尾的换行和多余文本
    """
    # Step 1: 去掉查询参数（? 后面的内容）
    if '?' in raw_path:
        raw_path = raw_path.split('?', 1)[0]

    # Step 2: 替换 ${...} 为 {id}
    path_with_braces = re.sub(r'\$\{[^}]*\}', '{id}', raw_path)
    
    # Step 3: 替换 <...> 为 {...}
    path_with_braces = re.sub(r'<([^>]*)>', r'{\1}', path_with_braces)
    
    # Step 4: 替换 ** 为 {id}
    path_with_braces = re.sub(r'\*\*', '{id}', path_with_braces)
    
    # Step 5: 移除路径末尾的换行和多余文本
    path_with_braces = re.sub(r'[\r\n].*$', '', path_with_braces)
    
    # Step 6: 移除末尾可能的逗号或其他符号
    path_with_braces = re.sub(r'[,}]*$', '', path_with_braces)
    
    # Step 7: 修复不完整的花括号
    path_with_braces = re.sub(r'\{id[^}]*$', '{id}', path_with_braces)

    return path_with_braces


def extract_and_normalize_apis(text):
    results = []

    # 匹配 type: "method", url: "path" 格式
    # 支持单引号、双引号
    pattern = r'type:\s*["\']([^"\']*)["\']\s*,\s*url:\s*["\']([^"\']*)["\']'
    matches = re.findall(pattern, text, re.IGNORECASE)

    for method, raw_path in matches:
        # 清理路径
        raw_path = raw_path.strip()
        
        # 如果路径不以 api/ 开头，则跳过
        if not raw_path.startswith("api/"):
            continue
            
        # 标准化路径
        clean_path = normalize_path(raw_path)

        # 确保路径以 / 开头
        if not clean_path.startswith("/"):
            clean_path = "/" + clean_path

        # 检查是否是API路径
        if not clean_path.startswith("/api/"):
            print(f"⚠️ 忽略非 API 路径：{clean_path}")
            continue

        # 提取服务名和剩余路径
        # /api/serviceName/path -> serviceName, /path
        api_parts = clean_path[5:]  # 去掉 "/api/" 前缀
        if "/" in api_parts:
            service, remaining_path = api_parts.split("/", 1)
            remaining_path = "/" + remaining_path
        else:
            service = api_parts
            remaining_path = "/"

        results.append((method.upper(), service, remaining_path))

    return results


def main():
    input_file = 'http_get_contents2.txt'
    output_file = 'normalized_apis2_fixed.xlsx'

    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()
    except FileNotFoundError:
        print(f"❌ 文件 '{input_file}' 未找到。")
        return

    calls = extract_and_normalize_apis(content)

    if not calls:
        print("⚠️ 未提取到有效 API。")
        return

    df = pd.DataFrame(calls, columns=['请求方法', '服务名称', '请求路径'])
    df = df.drop_duplicates().reset_index(drop=True)

    # 设置列宽自适应
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Normalized APIs')
        
        # 获取工作表对象
        worksheet = writer.sheets['Normalized APIs']
        
        # 设置列宽
        for i, col in enumerate(df.columns):
            # 计算列宽，最大限制为50个字符
            max_len = min(max(df[col].astype(str).map(len).max(), len(col)) + 2, 50)
            worksheet.set_column(i, i, max_len)

    print(f"✅ 成功生成 {len(df)} 条标准化 API 到 '{output_file}'")

    print("\n前10条预览：")
    print(df.head(10))


if __name__ == '__main__':
    main()