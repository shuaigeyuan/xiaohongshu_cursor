import pandas as pd
import requests
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

def download_image(url, save_path):
    """
    下载单张图片
    
    参数:
    url (str): 图片下载地址
    save_path (str): 图片保存路径
    
    返回:
    bool: 下载是否成功
    """
    try:
        # 设置请求头，模拟浏览器访问
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        # 发送请求下载图片
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        # 保存图片
        with open(save_path, 'wb') as f:
            f.write(response.content)
        
        return True
    except Exception as e:
        print(f"下载图片 {url} 失败: {e}")
        return False

def download_covers(input_file, output_dir):
    """
    根据条件下载小红书笔记封面图片
    
    参数:
    input_file (str): 输入的Excel文件路径
    output_dir (str): 图片保存目录
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(input_file)
        print(f"成功读取文件: {input_file}")
        print(f"总笔记数: {len(df)}")

        # 筛选符合条件的数据
        filtered_df = df[(df['粉丝数'] < 1000) & (df['互动量'] > 100)]
        print(f"符合条件的笔记数: {len(filtered_df)}")

        # 创建保存目录
        os.makedirs(output_dir, exist_ok=True)

        # 使用线程池并发下载图片
        successful_downloads = 0
        failed_downloads = 0

        with ThreadPoolExecutor(max_workers=10) as executor:
            # 存储下载任务
            futures = []
            
            for index, row in filtered_df.iterrows():
                cover_url = row['封面地址']
                
                # 生成文件名：索引_笔记标题.jpg
                filename = f"{index}_{row['笔记标题'][:20]}.jpg"
                filename = "".join(x for x in filename if x.isalnum() or x in ['_', '.'])  # 清理文件名
                save_path = os.path.join(output_dir, filename)
                
                # 提交下载任务
                future = executor.submit(download_image, cover_url, save_path)
                futures.append((future, save_path))

            # 等待所有任务完成并统计结果
            for future, save_path in futures:
                result = future.result()
                if result:
                    successful_downloads += 1
                else:
                    failed_downloads += 1
                    # 删除下载失败的文件
                    if os.path.exists(save_path):
                        os.remove(save_path)

        # 打印下载统计信息
        print(f"\n下载统计:")
        print(f"成功下载: {successful_downloads} 张图片")
        print(f"下载失败: {failed_downloads} 张图片")

    except Exception as e:
        print(f"处理过程中发生错误: {e}")

# 主程序入口
if __name__ == "__main__":
    # 输入文件路径
    input_file = r'C:\Users\yuanhua\Desktop\cursor\小红书抓取爆款文案\xiaohongshu.xlsx'
    
    # 输出目录
    output_dir = r'C:\Users\yuanhua\Desktop\cursor\小红书抓取爆款文案\xiaohongshu_covers'
    
    # 执行下载
    download_covers(input_file, output_dir) 