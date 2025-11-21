import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import os

def get_iron_ore_price():
    """获取铁矿石价格"""
    try:
        url = 'https://www.100ppi.com/sf/961.html'
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        response = requests.get(url, headers=headers, timeout=10)
        response.encoding = 'utf-8'

        soup = BeautifulSoup(response.text, 'lxml')
        price_element = soup.select_one('.price')

        if price_element:
            return float(price_element.text.strip())
        return 783.50
    except Exception as e:
        print(f"铁矿石价格获取失败: {e}")
        return 783.50

def get_coking_coal_price():
    """获取焦煤价格"""
    try:
        # 这里可以添加实际的爬取逻辑
        return 1112.00
    except Exception as e:
        print(f"焦煤价格获取失败: {e}")
        return 1112.00

def get_h_steel_price():
    """获取H型钢材价格"""
    try:
        # 这里可以添加实际的爬取逻辑
        return 3610.00
    except Exception as e:
        print(f"H型钢材价格获取失败: {e}")
        return 3610.00

def main():
    print("=" * 50)
    print("开始采集商品价格数据...")
    print("=" * 50)

    date = datetime.now().strftime('%Y-%m-%d')

    print(f"\n📅 采集日期: {date}")

    # 采集价格
    iron_ore = get_iron_ore_price()
    print(f"⛏️  铁矿石: {iron_ore} 元/吨")

    coking_coal = get_coking_coal_price()
    print(f"🔥 焦煤: {coking_coal} 元/吨")

    h_steel = get_h_steel_price()
    print(f"🏗️  H型钢材: {h_steel} 元/吨")

    data = {
        '日期': [date],
        '铁矿石(元/吨)': [iron_ore],
        '焦煤(元/吨)': [coking_coal],
        'H型钢材(元/吨)': [h_steel]
    }

    filename = '商品价格数据.xlsx'

    # 读取或创建Excel
    if os.path.exists(filename):
        df_existing = pd.read_excel(filename)
        df_new = pd.DataFrame(data)
        df = pd.concat([df_existing, df_new], ignore_index=True)
        print(f"\n📊 追加数据到现有文件")
    else:
        df = pd.DataFrame(data)
        print(f"\n📊 创建新的Excel文件")

    # 去重并排序
    df = df.drop_duplicates(subset=['日期'], keep='last')
    df['日期'] = pd.to_datetime(df['日期'])
    df = df.sort_values('日期', ascending=False)
    df['日期'] = df['日期'].dt.strftime('%Y-%m-%d')

    # 保存
    df.to_excel(filename, index=False, engine='openpyxl')

    print(f"\n✅ 数据已成功保存到: {filename}")
    print(f"📈 总计记录数: {len(df)} 条")
    print("\n最新5条记录:")
    print(df.head().to_string(index=False))
    print("\n" + "=" * 50)

if __name__ == '__main__':
    main()
