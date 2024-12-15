import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import logging
import openpyxl

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s')

def process_xiaohongshu_notes(input_file, output_file):
    """
    处理小红书笔记数据的主函数
    
    参数:
    input_file (str): 输入的Excel文件路径
    output_file (str): 输出的Excel文件路径
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(input_file)
        logging.info(f"成功读取文件: {input_file}")
        logging.info(f"总笔记数: {len(df)}")

        # 定义请求头，包含必要的Cookie
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Cookie': 'abRequestId=923a3da1-c72a-518d-8dad-4457405ac299; a1=19313c5e3adjb7jwqy8ezsgtbi7d2nedmvfsdaisa50000120616; webId=fffd5108c5a564e838b8eacd094ca714; gid=yjqyqSK8fJ68yjqyqS2dqh030fVDWVEAvYduiM0Mk9D3WM28fJIdff888yJ8KyK8W2YJWi00; x-user-id-creator.xiaohongshu.com=5e06dfca0000000001006edc; customerClientId=491415899589480; webBuild=4.47.1; acw_tc=0a4a1d1e17341918816378573e0241d0ab4b575ae6fdf231e07c02e6ab1c63; web_session=0400698f729c82563e8ac7e568354bbfd94597; xsecappid=xhs-pc-web; websectiga=7750c37de43b7be9de8ed9ff8ea0e576519e8cd2157322eb972ecb429a7735d4; sec_poison_id=27e9a239-869f-4742-9cdc-143b9b13940a; unread={%22ub%22:%22673c47f80000000008004383%22%2C%22ue%22:%22674f2e9a000000000202a0a6%22%2C%22uc%22:31}'  # 替换为实际的Cookie
        }

        # 存储处理后的数据
        details = []
        hashtags = []

        # 遍历笔记链接
        for index, row in df.iterrows():
            note_url = row.iloc[0]  # 获取笔记链接
            logging.info(f"正在处理第 {index+1} 个笔记: {note_url}")

            try:
                # 发送请求获取网页内容
                response = requests.get(note_url, headers=headers, timeout=10)
                response.raise_for_status()  # 检查请求是否成功
                
                # 使用BeautifulSoup解析网页
                soup = BeautifulSoup(response.text, 'html.parser')

                # 提取笔记详情
                detail_elem = soup.find(id='detail-desc')
                detail_text = detail_elem.get_text(strip=True) if detail_elem else ''
                details.append(detail_text)

                # 提取笔记话题（可能有多个）
                hashtag_elems = soup.find_all(id='hash-tag')
                hashtag_texts = [tag.get_text(strip=True) for tag in hashtag_elems]
                hashtags.append(' | '.join(hashtag_texts))  # 多个话题用 | 分隔

                # 避免频繁请求
                time.sleep(1)

            except requests.RequestException as e:
                logging.error(f"请求笔记 {note_url} 失败: {e}")
                details.append('')
                hashtags.append('')

        # 将新列添加到DataFrame
        df['笔记详情'] = details
        df['笔记话题'] = hashtags

        # 保存到新的Excel文件
        df.to_excel(output_file, index=False)
        logging.info(f"处理完成，结果已保存到: {output_file}")

    except Exception as e:
        logging.error(f"处理过程中发生错误: {e}")

# 主程序入口
if __name__ == "__main__":
    input_file = r'C:\Users\yuanhua\Desktop\cursor\小红书抓取爆款文案\小学资料.xlsx'
    output_file = r'C:\Users\yuanhua\Desktop\cursor\小红书抓取爆款文案\xiaohongshu.xlsx'
    
    process_xiaohongshu_notes(input_file, output_file)