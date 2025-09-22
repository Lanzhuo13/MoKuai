import pandas as pd
import re
import json
import os
import sys
from atexit import register
import logging

# ===================== 配置区（可自定义） =====================
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))  # 项目根目录
CONFIG_FILE = os.path.join(PROJECT_ROOT, "配置中心.json")    # 配置文件路径
LOG_FILE = os.path.join(PROJECT_ROOT, "处理日志.log")       # 日志文件路径

# 默认配置（当配置文件不存在或字段缺失时使用）
DEFAULT_CONFIG = {
    "color_merging": {
        "_comment": "颜色合并规则：键=原始颜色，值=合并后颜色",
        "粉红色": "粉色",
        "浅灰": "浅灰色",
        "深黑": "黑色",
        "米白": "米色",
        "俄罗斯蓝": "蓝色"
    },
    "pattern_merging": {
        "_comment": "图案合并规则：键=原始图案，值=合并后图案",
        "彩虹马刺绣": "小马刺绣",
        "战马刺绣": "战马刺绣",
        "格子纹": "格子",
        "条纹图案": "条纹",
        "12543": "MA"
    },
    "color_dictionary": {
        "_comment": "基础颜色字典：新颜色会自动去重追加",
        "values": [
            "黑色", "白色", "红色", "蓝色", "绿色", "黄色", "粉色", "紫色",
            "灰色", "棕色", "橙色", "金色", "银色", "青色", "咖啡色", "米色",
            "卡其色", "藏青色", "军绿色", "杏色", "透明", "花色", "椰果白",
            "深蓝", "浅绿", "粉红", "紫红", "雾霾蓝", "豆绿色", "深灰色",
            "浅灰", "卡其"
        ]
    }
}

# 初始化日志
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)
# ==============================================================

# ---------------------- 配置文件初始化 ----------------------
os.makedirs(PROJECT_ROOT, exist_ok=True)  # 确保项目目录存在

def load_config():
    """加载配置文件（包含颜色合并规则、图案合并规则和基础颜色字典）"""
    try:
        # 检查配置文件是否存在
        if not os.path.exists(CONFIG_FILE):
            logging.info(f"配置文件'{CONFIG_FILE}'不存在，已创建默认配置")
            save_config(DEFAULT_CONFIG)
            return DEFAULT_CONFIG.copy()
        
        # 读取配置文件
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
        
        # 验证配置完整性
        config = validate_config(config)
        
        return config
    
    except json.JSONDecodeError:
        logging.error(f"配置文件格式错误，已重置为默认配置")
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()
    except Exception as e:
        logging.error(f"配置加载失败：{str(e)}，已重置为默认配置")
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()

def validate_config(config):
    """验证配置完整性，缺失字段时使用默认值"""
    # 颜色合并规则
    if "color_merging" not in config:
        config["color_merging"] = DEFAULT_CONFIG["color_merging"]
        logging.warning("配置文件中缺少color_merging字段，已使用默认值")
    
    # 图案合并规则
    if "pattern_merging" not in config:
        config["pattern_merging"] = DEFAULT_CONFIG["pattern_merging"]
        logging.warning("配置文件中缺少pattern_merging字段，已使用默认值")
    
    # 基础颜色字典
    if "color_dictionary" not in config:
        config["color_dictionary"] = DEFAULT_CONFIG["color_dictionary"]
        logging.warning("配置文件中缺少color_dictionary字段，已使用默认值")
    elif "values" not in config["color_dictionary"]:
        config["color_dictionary"]["values"] = DEFAULT_CONFIG["color_dictionary"]["values"]
        logging.warning("配置文件中color_dictionary缺少values字段，已使用默认值")
    
    return config

def save_config(config):
    """保存配置到文件"""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        logging.info(f"配置已保存至：{CONFIG_FILE}")
    except Exception as e:
        logging.error(f"保存配置失败：{str(e)}")

# 加载全局配置
global_config = load_config()
color_merging = global_config["color_merging"]
pattern_merging = global_config["pattern_merging"]
color_dictionary = global_config["color_dictionary"]["values"]
new_colors = set()  # 存储新发现的颜色

# 注册退出时自动保存配置
register(lambda: save_config({
    "color_merging": color_merging,
    "pattern_merging": pattern_merging,
    "color_dictionary": {
        "values": list(set(color_dictionary + list(new_colors)))
    }
}))


# ---------------------- 核心解析函数 ----------------------
def find_color_in_text(text):
    """在文本开头查找字典中的颜色（最长匹配）"""
    global color_dictionary, new_colors
    
    # 获取颜色键并排序（按长度降序）
    color_keys = sorted(color_dictionary, key=lambda x: len(x), reverse=True)
    
    # 优先匹配基础颜色字典（从文本开头匹配）
    for key in color_keys:
        if text.startswith(key):
            # 找到颜色后，从文本中移除该颜色
            remaining = text[len(key):].strip()
            return key, remaining
    
    # 未匹配到任何颜色
    return None, text

def extract_pattern_from_brackets(text):
    """从括号部分提取图案（去除括号）"""
    # 匹配括号内容（支持中英文括号）
    bracket_match = re.search(r'[\(（]([^\)）]+)[\)）]', text)
    if bracket_match:
        return bracket_match.group(1).strip()
    return None

def apply_merging_rules(color, pattern):
    """应用合并规则进行标准化"""
    # 应用颜色合并规则
    merged_color = color_merging.get(color, color)
    
    # 应用图案合并规则
    merged_pattern = pattern_merging.get(pattern, pattern)
    
    return merged_color, merged_pattern

def process_segment(segment):
    """单条数据完整处理流程（优化版）"""
    try:
        segment = str(segment).strip()  # 转换为字符串并去除空格
        if not segment:
            return "无颜色", "无图案"
        
        # 1. 尝试提取括号内容作为图案
        pattern = extract_pattern_from_brackets(segment)
        if pattern:
            # 移除括号部分
            text_without_brackets = re.sub(r'[\(（][^\)）]+[\)）]', '', segment).strip()
            
            # 在剩余文本中查找颜色（从开头匹配）
            color, remaining = find_color_in_text(text_without_brackets)
            
            if color:
                # 找到颜色，图案只保留括号内容（清空剩余文本）
                # 不将剩余文本添加到图案
                pass
            else:
                # 未找到颜色，整个剩余文本作为颜色
                color = text_without_brackets
        else:
            # 2. 无括号时，在整个文本中查找颜色（从开头匹配）
            color, remaining = find_color_in_text(segment)
            
            if color:
                # 找到颜色，剩余文本作为图案
                pattern = remaining if remaining else "无图案"
            else:
                # 未找到颜色，整个文本作为颜色
                color = segment
                pattern = "无图案"
        
        # 3. 处理未找到颜色的情况
        if not color:
            color = "无颜色"
        
        # 4. 处理未找到图案的情况
        if not pattern:
            pattern = "无图案"
        
        # 5. 应用合并规则进行标准化
        color, pattern = apply_merging_rules(color, pattern)
        
        # 6. 记录新颜色
        if color not in color_dictionary and color != "无颜色":
            new_colors.add(color)
            logging.info(f"发现新颜色: {color}，已添加到颜色字典")
        
        return color, pattern
    
    except Exception as e:
        logging.error(f"处理异常：{str(e)} → 输入文本：{segment}")
        return "无颜色", "无图案"


# ---------------------- Excel处理函数 ----------------------
def process_excel(input_path, output_path):
    """处理Excel文件（读取→解析→保存）"""
    try:
        # 读取Excel（强制转换为字符串，避免类型问题）
        df = pd.read_excel(input_path, sheet_name="Sheet1", dtype=str)
        logging.info(f"成功读取Excel文件: {input_path}")
    except Exception as e:
        logging.error(f"Excel读取失败：{str(e)}")
        print(f"Excel读取失败：{str(e)}")
        return
    
    # 初始化结果列
    df["处理后颜色"] = ""
    df["处理后图案"] = ""
    
    # 逐行处理
    processed_count = 0
    for idx, row in df.iterrows():
        original_text = row["前段字段"]
        
        # 处理空值（NaN或空字符串）
        if pd.isna(original_text) or not original_text.strip():
            df.at[idx, "处理后颜色"] = "无颜色"
            df.at[idx, "处理后图案"] = "无图案"
            continue
        
        # 解析颜色和图案
        color, pattern = process_segment(original_text)
        df.at[idx, "处理后颜色"] = color
        df.at[idx, "处理后图案"] = pattern
        processed_count += 1
    
    # 保存结果
    try:
        df.to_excel(
            output_path,
            index=False,
            sheet_name="处理结果",
            engine="openpyxl"
        )
        msg = f"处理完成！共处理{processed_count}条记录，结果已保存至：{output_path}"
        logging.info(msg)
        print(msg)
        
        # 记录新颜色发现情况
        if new_colors:
            logging.info(f"发现{len(new_colors)}个新颜色：{', '.join(new_colors)}")
            print(f"发现{len(new_colors)}个新颜色，已添加到颜色字典")

    except Exception as e:
        logging.error(f"Excel保存失败：{str(e)}")
        print(f"Excel保存失败：{str(e)}")


# ---------------------- 主程序入口 ----------------------
if __name__ == "__main__":
    # 输入输出路径（自动适配项目根目录）
    input_file = os.path.join(PROJECT_ROOT, "分割处理.xlsx")
    output_file = os.path.join(PROJECT_ROOT, "颜色图案处理.xlsx")
    
    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        error_msg = f"错误：输入文件'{input_file}'不存在！"
        logging.error(error_msg)
        print(error_msg)
        sys.exit(1)
    
    # 执行处理
    process_excel(input_file, output_file)