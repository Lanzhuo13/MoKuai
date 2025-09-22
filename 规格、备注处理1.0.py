import re
import pandas as pd
from pathlib import Path

def extract_spec_remark(back_text: str) -> tuple:
    """
    从后段字段提取规格和备注
    处理规则：
    1. 优先删除规格后相邻括号内容（含空格情况）
    2. 使用预定义规格模式匹配
    3. 动态分割剩余内容作为备注
    """
    # 预定义规格正则模式（可扩展）
    SPEC_PATTERNS = [
        r'均码', r'[X]{1,3}[L]', r'\d{2,3}[A-Z]?', 
        r'中国(码|号型[A-CY])', r'腰围\d+', r'通用码'
    ]
    
    # 1. 处理括号内容（含空格情况）
    bracket_pattern = r'^\s*([^\s\(\)]+?)\s*(?:\([^)]*\))\s*(.*)$'
    if match := re.match(bracket_pattern, back_text):
        spec = match.group(1)
        remark = match.group(2).strip()
        return spec, remark
    
    # 2. 正则优先模式匹配
    combined_pattern = r'^\s*(' + '|'.join(SPEC_PATTERNS) + r')(.*)$'
    if match := re.match(combined_pattern, back_text, re.IGNORECASE):
        spec = match.group(1).strip()
        remark = match.group(2).strip()
        return spec, remark
    
    # 3. 动态分隔检测
    # 智能识别数字/中文边界
    if match := re.match(r'^(\d+)([^\d].*)$', back_text):
        return match.group(1), match.group(2)
    
    # 4. 最后尝试空格分割
    parts = back_text.strip().split(maxsplit=1)
    spec = parts[0] if parts else ""
    remark = parts[1] if len(parts) > 1 else ""
    
    return spec, remark

def process_excel(input_path: str, output_path: str):
    """
    主处理函数：读取Excel，处理规格备注，保存结果
    """
    # 读取输入文件
    df = pd.read_excel(input_path)
    
    # 检查必要列是否存在
    if '后段字段' not in df.columns:
        raise ValueError("输入文件缺少'后段字段'列")
    
    # 处理每条记录
    results = []
    for _, row in df.iterrows():
        back_text = str(row['后段字段'])
        spec, remark = extract_spec_remark(back_text)
        results.append({'原始后段字段': back_text, '规格': spec, '备注': remark})
    
    # 创建结果DataFrame并保存
    result_df = pd.DataFrame(results)
    result_df.to_excel(output_path, index=False)
    return f"处理完成，结果已保存至: {output_path}"

# 文件路径配置
input_file = r"D:\桌面\流程代码\分割处理.xlsx"
output_file = r"D:\桌面\流程代码\规格、备注处理.xlsx"

# 执行处理
if __name__ == "__main__":
    print(process_excel(input_file, output_file))