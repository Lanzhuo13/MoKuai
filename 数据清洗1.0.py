import pandas as pd
import os
import argparse
import sys
from difflib import get_close_matches
from typing import Dict, List, Optional

def clean_excel_data(
    input_path: str = r"D:\桌面\流程代码\复合样本.xlsx",
    output_path: str = "解析文件.xlsx",
    valid_columns: Dict[str, List[str]] = None,
    alias_priority: bool = True,
    similarity_threshold: float = 0.8,
    sheet_name: str = 'Sheet1'
) -> None:
    """
    Excel数据清洗工具
    功能特性：
    - 自动删除「总数量」列
    - 别名优先匹配（支持多别名映射）
    - 模糊匹配容错（Levenshtein距离算法）
    - 冲突解决策略（优先级/相似度排序）
    - 默认路径配置（Windows系统）
    - 精准删除无效行（规格为空或包含汇总/合计）
    - 智能处理合并单元格（自动填充宝贝名称）

    参数说明：
    valid_columns: {
        "标准列名": ["别名1", "别名2"],
        "数量": ["数量", "件数", "总数量"]
    }
    alias_priority: 是否优先使用别名匹配（True）还是模糊匹配（False）
    similarity_threshold: 模糊匹配最低相似度阈值（0.0-1.0）
    """
    
    # 默认有效列配置
    default_valid_columns = {
        "宝贝名称": ["商品名称", "品名"],
        "宝贝规格": ["规格", "型号", "尺寸"],
        "数量": ["数量", "件数", "库存", "销量"],
    }
    
    valid_columns = valid_columns or default_valid_columns
    
    try:
        # 检查输入文件是否存在
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"输入文件未找到：{input_path}")
        
        print(f"正在处理文件: {input_path}")
        
        # 读取原始数据
        df = pd.read_excel(input_path, sheet_name=sheet_name, dtype=str)
        print(f"原始数据行数: {len(df)}")
        
        # 预处理列名
        original_columns = df.columns.tolist()
        cleaned_columns = [col.strip().replace(' ', '_').lower() for col in original_columns]
        
        # 构建列名映射
        column_mapping = {}
        for std_col, aliases in valid_columns.items():
            for alias in aliases:
                alias_key = alias.strip().replace(' ', '_').lower()
                column_mapping[alias_key] = std_col
            
            std_key = std_col.strip().replace(' ', '_').lower()
            if std_key not in column_mapping:
                column_mapping[std_key] = std_col
        
        # 执行列匹配
        matched_columns = []
        unmatched_columns = []
        for orig_col, clean_col in zip(original_columns, cleaned_columns):
            if clean_col in column_mapping:
                matched_columns.append((orig_col, column_mapping[clean_col]))
            else:
                unmatched_columns.append(orig_col)
        
        # 模糊匹配处理未匹配列
        if unmatched_columns and not alias_priority:
            for col in unmatched_columns:
                matches = get_close_matches(
                    col.lower(),
                    column_mapping.keys(),
                    n=1,
                    cutoff=similarity_threshold
                )
                
                if matches:
                    matched_columns.append((col, column_mapping[matches[0]]))
                else:
                    matched_columns.append((col, None))
        else:
            matched_columns += [(col, None) for col in unmatched_columns]
        
        # 构建列保留映射
        keep_columns = []
        column_aliases = {}
        for orig_col, mapped_col in matched_columns:
            if mapped_col:
                keep_columns.append(orig_col)
                column_aliases[orig_col] = mapped_col
        
        # 数据过滤与重命名
        df_cleaned = df[keep_columns].copy()
        df_cleaned.rename(columns=column_aliases, inplace=True)
        
        # === 新增处理逻辑：精准删除无效行 ===
        # 1. 删除宝贝规格为空的无效行
        if '宝贝规格' in df_cleaned.columns:
            # 创建空值检测掩码
            empty_mask = df_cleaned['宝贝规格'].isna() | (df_cleaned['宝贝规格'].str.strip() == '')
            print(f"检测到 {empty_mask.sum()} 行宝贝规格为空，将被删除")
            df_cleaned = df_cleaned[~empty_mask]
        
        # 2. 删除包含"汇总"、"合计"字样的行
        summary_keywords = ['汇总', '合计']
        summary_mask = pd.Series(False, index=df_cleaned.index)
        
        for col in df_cleaned.columns:
            if pd.api.types.is_string_dtype(df_cleaned[col]):
                # 创建每列的关键词检测掩码
                col_mask = df_cleaned[col].str.contains('|'.join(summary_keywords), na=False)
                summary_mask = summary_mask | col_mask
        
        print(f"检测到 {summary_mask.sum()} 行包含汇总/合计字样，将被删除")
        df_cleaned = df_cleaned[~summary_mask]
        
        # === 新增处理逻辑：智能处理合并单元格 ===
        # 3. 自动填充宝贝名称列（处理合并单元格）
        if '宝贝名称' in df_cleaned.columns:
            # 记录填充前的空值数量
            na_count_before = df_cleaned['宝贝名称'].isna().sum()
            
            # 向前填充空值
            df_cleaned['宝贝名称'] = df_cleaned['宝贝名称'].ffill()
            
            # 特殊处理：当宝贝名称为空但规格存在时
            if '宝贝规格' in df_cleaned.columns:
                # 创建有效数据掩码
                valid_mask = df_cleaned['宝贝规格'].notna() & (df_cleaned['宝贝规格'].str.strip() != '')
                
                # 填充剩余空值
                df_cleaned.loc[valid_mask, '宝贝名称'] = df_cleaned['宝贝名称'].where(
                    df_cleaned['宝贝名称'].notna(), '未命名商品'
                )
            
            # 记录填充后的空值数量
            na_count_after = df_cleaned['宝贝名称'].isna().sum()
            print(f"宝贝名称列填充完成：填充了 {na_count_before - na_count_after} 个空值")
        
        # === 原有处理逻辑 ===
        # 删除冗余列
        if '总数量' in df_cleaned.columns:
            df_cleaned.drop('总数量', axis=1, inplace=True, errors='ignore')
            print("已删除「总数量」列")
        
        # 补充缺失标准列
        for std_col in valid_columns.keys():
            if std_col not in df_cleaned.columns:
                df_cleaned[std_col] = None
        
        # 保存结果
        df_cleaned.to_excel(output_path, index=False, sheet_name='CleanedData')
        
        print(f"清洗完成！共处理 {len(df_cleaned)} 条数据")
        print(f"有效列：{', '.join(df_cleaned.columns)}")
        print(f"输出文件: {output_path}")
        
    except FileNotFoundError as e:
        print(f"错误：{str(e)}")
    except Exception as e:
        print(f"清洗失败：{str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # 检查是否提供了命令行参数
    if len(sys.argv) > 1:
        # 命令行参数解析
        parser = argparse.ArgumentParser(description="Excel数据清洗工具")
        parser.add_argument("--input", type=str, required=True, help="输入Excel文件路径")
        parser.add_argument("--output", type=str, default="cleaned_data.xlsx", help="输出文件路径")
        parser.add_argument("--alias_priority", action="store_true", help="启用别名优先匹配")
        parser.add_argument("--similarity", type=float, default=0.8, help="模糊匹配相似度阈值")
        args = parser.parse_args()
        
        # 执行清洗
        clean_excel_data(
            input_path=args.input,
            output_path=args.output,
            alias_priority=args.alias_priority,
            similarity_threshold=args.similarity
        )
    else:
        # 无命令行参数时使用默认路径
        print("未提供命令行参数，使用默认路径...")
        clean_excel_data()