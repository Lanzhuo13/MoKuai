import pandas as pd
import os
import logging
from typing import List, Tuple

# ------------------------------
# 全局配置与日志初始化（保持不变）
# ------------------------------
def init_logger(output_dir: str) -> logging.Logger:
    """初始化日志系统"""
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    
    # 日志格式
    formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # 文件 handler（保存到输出目录）
    file_handler = logging.FileHandler(os.path.join(output_dir, "字段分割.log"))
    file_handler.setFormatter(formatter)
    
    # 控制台 handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger


# ------------------------------
# 核心分割逻辑（保持不变）
# ------------------------------
def split_outside_parentheses(
    text: str,
    separators: List[str] = [',', ' ']  # 默认分隔符：逗号、空格
) -> Tuple[str, str]:
    if not text:
        return "", ""
    
    bracket_depth = 0  # 括号嵌套深度（0表示不在括号内）
    first_valid_sep_pos = -1  # 记录首个有效分隔符位置
    
    # 遍历每个字符，追踪括号状态
    for idx, char in enumerate(text):
        if char == '(':
            bracket_depth += 1  # 进入括号，深度+1
        elif char == ')':
            bracket_depth = max(0, bracket_depth - 1)  # 退出括号，深度-1（防止负数）
        elif bracket_depth == 0 and char in separators:
            first_valid_sep_pos = idx  # 找到括号外的首个分隔符，记录位置并终止遍历
            break
    
    # 未找到有效分隔符，返回原文本和空字符串
    if first_valid_sep_pos == -1:
        return text, ""
    
    # 分割文本并清理首尾多余分隔符
    front_part = text[:first_valid_sep_pos].rstrip(''.join(separators))
    back_part = text[first_valid_sep_pos:].lstrip(''.join(separators))
    
    # 处理分割后为空的极端情况
    if not front_part or not back_part:
        return text, ""
    
    return front_part, back_part


# ------------------------------
# Excel批量处理函数（调整：接收logger参数）
# ------------------------------
def process_field_split(
    input_path: str = r"D:\桌面\流程代码\分割数据源.xlsx",
    output_path: str = None,
    text_column: str = "商品描述",  # 需要分割的原始文本列名
    front_column: str = "前段字段",  # 分割后的前段字段名
    back_column: str = "后段字段",  # 分割后的后段字段名
    separators: List[str] = [',', ' '],  # 自定义分隔符列表
    logger: logging.Logger = None  # 新增：接收外部logger实例
) -> None:
    """
    Excel字段分割主流程：读取→分割→保存
    :param logger: 外部传入的logger实例（用于记录日志）
    """
    # 1. 初始化输出路径与日志（若未传入logger，则内部初始化）
    output_dir = os.path.dirname(input_path)
    output_path = output_path or os.path.join(output_dir, "分割处理.xlsx")
    
    # 如果未传入logger，内部初始化（保证函数独立性）
    if not logger:
        logger = init_logger(output_dir)
    
    logger.info(f"=== 开始处理字段分割 ===")
    logger.info(f"输入文件：{input_path}")
    logger.info(f"输出文件：{output_path}")
    logger.info(f"文本列：{text_column} | 分隔符：{separators}")


    # 2. 读取Excel数据（保持不变）
    try:
        df = pd.read_excel(input_path)
        logger.info(f"成功读取Excel，共{len(df)}条数据")
    except Exception as e:
        logger.error(f"读取Excel失败：{str(e)}", exc_info=True)
        raise


    # 3. 校验文本列是否存在（保持不变）
    if text_column not in df.columns:
        error_msg = f"文本列'{text_column}'不存在于Excel中"
        logger.error(error_msg)
        raise ValueError(error_msg)


    # 4. 逐行处理字段分割（保持不变）
    front_results = []
    back_results = []
    error_records = []  # 记录错误行信息（行号、文本、错误原因）

    for idx, row in df.iterrows():
        text = str(row[text_column]).strip()  # 处理空值或非字符串
        row_num = idx + 2  # Excel行号（header占1行，索引从0开始）
        
        try:
            front, back = split_outside_parentheses(text, separators)
            front_results.append(front)
            back_results.append(back)
        except Exception as e:
            front_results.append("")
            back_results.append("")
            error_records.append((row_num, text, str(e)))
            logger.warning(f"行{row_num}处理失败：{text} → {str(e)}")


    # 5. 将结果写入DataFrame（保持不变）
    df[front_column] = front_results
    df[back_column] = back_results


    # 6. 保存结果到Excel（保持不变）
    try:
        df.to_excel(output_path, index=False)
        logger.info(f"结果保存成功，共{len(df)}条记录")
        
        # 记录错误信息（如果有）
        if error_records:
            logger.warning(f"共{len(error_records)}条错误：")
            for row_num, text, err in error_records:
                logger.warning(f"行{row_num} | 文本：{text[:50]}... | 错误：{err}")
    except Exception as e:
        logger.error(f"保存Excel失败：{str(e)}", exc_info=True)
        raise


# ------------------------------
# 主程序入口（修正：初始化全局logger）
# ------------------------------
if __name__ == "__main__":
    try:
        # 1. 初始化全局logger（关键修正：让主程序能访问logger）
        output_dir = os.path.dirname(r"D:\桌面\流程代码\分割数据源.xlsx")
        logger = init_logger(output_dir)
        
        # 2. 执行字段分割（传递logger给函数）
        process_field_split(
            input_path=r"D:\桌面\流程代码\分割数据源.xlsx",  # 明确指定输入路径（可选）
            text_column="宝贝规格",  # 替换为您的文本列名
            separators=[',', ' ', ';'],  # 可选：扩展分隔符（如分号）
            logger=logger  # 传递logger实例
        )
        
        # 3. 处理完成日志（现在logger有定义了）
        logger.info("=== 处理完成 ===")
        
    except Exception as e:
        # 4. 异常处理（使用全局logger）
        logger.critical(f"程序终止：{str(e)}", exc_info=True)