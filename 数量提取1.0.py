import pandas as pd
import numpy as np
import re
from typing import List, Tuple, Optional

class QuantityExtractionError(Exception):
    """自定义数量提取异常类"""
    pass

# 全局配置中心
class ConfigCenter:
    """配置中心类"""
    PRIORITY = "independent"  # 可选值: "independent"（独立列优先）或 "spec"（规格文本优先）
    ALLOW_CONFLICT = True     # 是否允许数量冲突
    LOG_CONFLICTS = True      # 是否记录冲突日志

# 全局分隔符配置
BASE_SEPARATORS = [':', '*']
EXTENDED_SEPARATORS = []

def register_quantity_separators(new_separators: list):
    """注册新的数量分隔符"""
    global EXTENDED_SEPARATORS
    # 去重并保留唯一值
    EXTENDED_SEPARATORS = list(set(EXTENDED_SEPARATORS + new_separators))

def extract_qty_from_spec(spec_text: str) -> Tuple[Optional[float], str]:
    """
    从规格文本中提取数量（增强版）
    返回: (提取的数量, 清理后的规格文本)
    """
    if pd.isna(spec_text) or not isinstance(spec_text, str):
        return None, spec_text
    
    # 合并所有分隔符
    all_seps = BASE_SEPARATORS + EXTENDED_SEPARATORS
    
    # 从后向前查找第一个分隔符（包括空格变体）
    found_index = -1
    found_sep = None
    
    # 检查所有分隔符及其空格变体
    for sep in all_seps:
        # 检查分隔符本身
        index = spec_text.rfind(sep)
        if index > found_index:
            found_index = index
            found_sep = sep
        
        # 检查分隔符前后带空格的情况
        for space in [" ", "  ", "   "]:
            # 分隔符前加空格
            spaced_sep = space + sep
            index = spec_text.rfind(spaced_sep)
            if index > found_index:
                found_index = index
                found_sep = spaced_sep
            
            # 分隔符后加空格
            spaced_sep = sep + space
            index = spec_text.rfind(spaced_sep)
            if index > found_index:
                found_index = index
                found_sep = spaced_sep
            
            # 分隔符前后都加空格
            spaced_sep = space + sep + space
            index = spec_text.rfind(spaced_sep)
            if index > found_index:
                found_index = index
                found_sep = spaced_sep
    
    # 未找到分隔符
    if found_index == -1:
        return None, spec_text
    
    # 计算数字部分起始位置
    num_start = found_index + len(found_sep)
    
    # 提取数字部分
    num_part = spec_text[num_start:].strip()
    num_match = re.search(r'^\d+', num_part)
    
    if num_match:
        quantity = float(num_match.group(0))
        clean_text = spec_text[:found_index].strip()
        return quantity, clean_text
    
    return None, spec_text

def extract_quantity(
    file_path: str = r"D:\桌面\流程代码\解析文件.xlsx",
    output_path: str = r"D:\桌面\流程代码\数量提取.xlsx",
    secondary_columns: List[str] = ["件数", "宝贝数量", "商品数量"],
    custom_secondary: List[str] = None,
    spec_col: str = "宝贝规格"  # 规格文本列名
) -> None:
    """
    数量提取核心处理函数（修复独立列数量问题）
    :param file_path: 输入Excel文件路径
    :param output_path: 输出Excel文件路径
    :param secondary_columns: 预定义次级候选列名列表
    :param custom_secondary: 用户自定义次级候选列名列表
    :param spec_col: 规格文本列名
    """
    # 加载数据
    df = pd.read_excel(file_path)
    
    # 主列检查
    primary_col = "数量"
    if primary_col in df.columns:
        quantity_col = primary_col
    else:
        # 合并次级候选列（优先自定义）
        check_columns = secondary_columns.copy()
        if custom_secondary:
            check_columns.extend(custom_secondary)
        
        # 查找匹配的次级列
        matched_cols = [col for col in check_columns if col in df.columns]
        
        if len(matched_cols) > 1:
            raise QuantityExtractionError("发现多个候选数量列，请检查数据格式！")
        elif len(matched_cols) == 1:
            quantity_col = matched_cols[0]
        else:
            raise QuantityExtractionError("未找到有效数量列，请检查数据格式！")
    
    # === 核心处理逻辑 ===
    # 1. 从独立列提取数量（优先）
    df['独立列数量'] = pd.to_numeric(df[quantity_col], errors='coerce')
    
    # 2. 初始化最终数量为独立列数量（关键修复）
    df['最终数量'] = df['独立列数量']
    
    # 3. 从规格文本提取数量（条件性）
    if spec_col in df.columns:
        # 应用提取函数
        df[['规格提取数量', '清理后规格']] = df.apply(
            lambda row: pd.Series(extract_qty_from_spec(row[spec_col])),
            axis=1
        )
        
        # 只有当独立列数量为空时才使用规格提取数量
        mask = df['独立列数量'].isna()
        df.loc[mask, '最终数量'] = df.loc[mask, '规格提取数量']
    
    # 4. 冲突检测（如果配置允许）
    if ConfigCenter.ALLOW_CONFLICT and spec_col in df.columns:
        conflict_mask = (
            df['规格提取数量'].notna() & 
            df['独立列数量'].notna() &
            (df['规格提取数量'] != df['独立列数量'])
        )
        df['数量冲突'] = conflict_mask
        
        # 记录冲突详情（如果配置允许）
        if ConfigCenter.LOG_CONFLICTS:
            df['冲突详情'] = np.where(
                conflict_mask,
                f"规格提取值:{df['规格提取数量']} vs 独立列值:{df['独立列数量']}",
                ""
            )
    
    # 5. 空值处理：NaN或None转为0
    df['最终数量'] = df['最终数量'].fillna(0)
    
    # 6. 保存结果
    df.to_excel(output_path, index=False)
    print(f"数量提取完成，结果已保存至：{output_path}")
    
    # 7. 打印冲突报告（如果配置允许）
    if (ConfigCenter.ALLOW_CONFLICT and 
        ConfigCenter.LOG_CONFLICTS and 
        '数量冲突' in df.columns and 
        df['数量冲突'].any()):
        print("\n=== 数量冲突警告 ===")
        for idx, row in df[df['数量冲突']].iterrows():
            print(f"行 {idx+2}: {row['冲突详情']} | 原始规格: {row.get(spec_col, '')}")

if __name__ == "__main__":
    try:
        # 注册自定义分隔符（可选）
        register_quantity_separators(["/", "#"])
        
        # 执行主处理
        extract_quantity(
            custom_secondary=["单品数量", "销售数量"]  # 用户可自定义扩展
        )
    except QuantityExtractionError as e:
        print(f"处理失败：{str(e)}")
    except Exception as e:
        print(f"未知错误：{str(e)}")