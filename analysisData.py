import pandas as pd
import jieba
from collections import Counter
import re

def preprocess_title(title):
    """
    预处理标题：去除特殊字符，保留中文和数字
    """
    # 使用正则表达式保留中文和数字
    title = re.sub(r'[^\u4e00-\u9fff0-9]+', '', str(title))
    return title

def extract_keywords(title):
    """
    使用结巴分词提取关键词
    """
    # 去除停用词
    stop_words = {'的', '了', '在', '是', '和', '与', '或', '我', '你', '他', '她', '它'}
    
    # 分词
    words = jieba.cut(title)
    
    # 过滤停用词和单字
    keywords = [word for word in words if word not in stop_words and len(word) > 1]
    
    return keywords

def analyze_titles(input_file, output_file):
    """
    分析小红书笔记标题
    
    参数:
    input_file (str): 输入的Excel文件路径
    output_file (str): 输出的Excel文件路径
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(input_file)
        print(f"成功读取文件: {input_file}")
        print(f"总笔记数: {len(df)}")

        # 预处��和关键词提取
        all_keywords = []
        title_mapping = {}

        for title in df['笔记标题']:
            # 预处理标题
            clean_title = preprocess_title(title)
            
            # 提取关键词
            keywords = extract_keywords(clean_title)
            
            for keyword in keywords:
                all_keywords.append(keyword)
                if keyword not in title_mapping:
                    title_mapping[keyword] = set()
                title_mapping[keyword].add(title)

        # 统计关键词频率
        keyword_counter = Counter(all_keywords)
        
        # 按出现次数倒序排列
        sorted_keywords = sorted(keyword_counter.items(), key=lambda x: x[1], reverse=True)

        # 准备输出数据
        output_data = []
        for keyword, count in sorted_keywords:
            # 获取包含该关键词的所有标题
            related_titles = list(title_mapping[keyword])
            
            output_data.append({
                '关键词': keyword,
                '出现次数': count,
                '相关笔记标题': ' | '.join(related_titles[:10])  # 限制显示前10个标题
            })

        # 创建输出DataFrame
        output_df = pd.DataFrame(output_data)
        
        # 保存到新的Excel文件
        output_df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"分析结果已保存到: {output_file}")
        
        # 打印前20个关键词
        print("\n前20个高频关键词：")
        for keyword, count in sorted_keywords[:20]:
            print(f"{keyword}: {count}次")

    except Exception as e:
        print(f"处理过程中发生错误: {e}")

# 主程序入口
if __name__ == "__main__":
    # 输入文件路径
    input_file = r'C:\Users\yuanhua\Desktop\cursor\小红书抓取爆款文案\xiaohongshu.xlsx'
    
    # 输出文件路径
    output_file = r'C:\Users\yuanhua\Desktop\cursor\小红书抓取爆款文案\xiaohongshu_title_analysis.xlsx'
    
    # 执行分析
    analyze_titles(input_file, output_file)
