import pandas as pd
from collections import Counter

def analyze_hashtags(input_file, output_file):
    """
    分析小红书笔记的话题标签
    
    参数:
    input_file (str): 输入的Excel文件路径
    output_file (str): 输出的txt文件路径
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(input_file)
        print(f"成功读取文件: {input_file}")
        print(f"总笔记数: {len(df)}")

        # 存储所有话题
        all_hashtags = []

        # 遍历话题标签列
        for hashtag_str in df['笔记话题']:
            # 跳过空值
            if pd.isna(hashtag_str):
                continue
            
            # 按 | 分割话题
            hashtags = hashtag_str.split(' | ')
            
            # 去除 # 符号并清理空白
            hashtags = [tag.replace('#', '').strip() for tag in hashtags]
            
            # 添加到总列表
            all_hashtags.extend(hashtags)

        # 统计话题频率
        hashtag_counter = Counter(all_hashtags)
        
        # 按出现次数倒序排列，取前50个
        top_50_hashtags = hashtag_counter.most_common(50)

        # 写入txt文件
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("话题\t出现次数\n")
            f.write("-" * 20 + "\n")
            
            for hashtag, count in top_50_hashtags:
                f.write(f"{hashtag}\t{count}\n")

        # 打印结果
        print(f"\n分析结果已保存到: {output_file}")
        print("\n前10个高频话题：")
        for hashtag, count in top_50_hashtags[:10]:
            print(f"{hashtag}: {count}次")

    except Exception as e:
        print(f"处理过程中发生错误: {e}")

# 主程序入口
if __name__ == "__main__":
    # 输入文件路径
    input_file = r'C:\Users\yuanhua\Desktop\cursor\小红书抓取爆款文案\xiaohongshu.xlsx'
    
    # 输出文件路径
    output_file = r'C:\Users\yuanhua\Desktop\cursor\小红书抓取爆款文案\top_50_hashtags.txt'
    
    # 执行分析
    analyze_hashtags(input_file, output_file) 