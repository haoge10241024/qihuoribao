import os
import akshare as ak
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.pyplot as plt
import mplfinance as mpf

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

def get_market_trend_data(symbol, custom_date):
    try:
        today = custom_date
        yesterday = today - timedelta(days=1)
        start_time = yesterday.strftime('%Y-%m-%d') + ' 21:00:00'
        end_time = today.strftime('%Y-%m-%d') + ' 23:00:00'
        df = ak.futures_zh_minute_sina(symbol=symbol, period="1")
        df['datetime'] = pd.to_datetime(df['datetime'])
        filtered_data = df[(df['datetime'] >= start_time) & (df['datetime'] <= end_time)]
        if filtered_data.empty:
            raise ValueError("市场数据为空")

        day_open_price = filtered_data.iloc[0]['open']
        day_close_price = filtered_data[(filtered_data['datetime'] <= today.strftime('%Y-%m-%d') + ' 15:00:00')].iloc[-1]['close']

        high_price = filtered_data['high'].max()
        low_price = filtered_data['low'].min()
        price_change = day_close_price - day_open_price
        price_change_percentage = (price_change / day_open_price) * 100
        trend = "上涨" if price_change > 0 else "下跌" if price_change < 0 else "持平"
        day_description = (
            f"{custom_date.strftime('%Y-%m-%d')}日{symbol}主力合约开盘价为{day_open_price}元/吨，最高价为{high_price}元/吨，"
            f"最低价为{low_price}元/吨，收盘价为{day_close_price}元/吨，较前一日{trend}了"
            f"{abs(price_change):.2f}元/吨，涨跌幅为{price_change_percentage:.2f}%。"
        )

        night_start_time = today.strftime('%Y-%m-%d') + ' 21:00:00'
        night_end_time = (today + timedelta(days=1)).strftime('%Y-%m-%d') + ' 01:00:00'
        night_data = df[(df['datetime'] >= night_start_time) & (df['datetime'] <= night_end_time)]
        if night_data.empty:
            raise ValueError("夜盘数据为空")
        night_open_price = night_data.iloc[0]['open']
        night_close_price = night_data.iloc[-1]['close']
        night_price_change = night_close_price - night_open_price
        night_price_change_percentage = (night_price_change / night_open_price) * 100
        night_trend = "上涨" if night_price_change > 0 else "下跌" if night_price_change < 0 else "持平"
        night_description = (
            f"夜盘走势：开盘价为{night_open_price}元/吨，收盘价为{night_close_price}元/吨，较开盘{night_trend}了"
            f"{abs(night_price_change):.2f}元/吨，涨跌幅为{night_price_change_percentage:.2f}%。"
        )

        return day_description, night_description, filtered_data
    except Exception as e:
        print(f"获取市场走势数据失败: {e}")
        return "", "", pd.DataFrame()

def create_k_line_chart(data, symbol, folder_path):
    if data.empty:
        return None
    data.set_index('datetime', inplace=True)
    data = data[['open', 'high', 'low', 'close']]
    data.columns = ['Open', 'High', 'Low', 'Close']
    fig, ax = plt.subplots(figsize=(10, 6))
    mpf.plot(data, type='candle', style='charles', ax=ax)
    ax.set_title(f'{symbol} 当日K线图')
    plt.ylabel('')
    plt.xlabel('')
    plt.savefig(os.path.join(folder_path, 'k_line_chart.png'), dpi=300)
    plt.close(fig)
    return os.path.join(folder_path, 'k_line_chart.png')

def get_news_data(symbol, custom_date):
    try:
        news_mapping = {
            '铜': '铜',
            '铝': '铝',
            '铅': '铅',
            '锌': '锌',
            '镍': '镍',
            '锡': '锡'
        }
        chinese_symbol = news_mapping.get(symbol, '铜')
        df = ak.futures_news_shmet(symbol=chinese_symbol)
        start_time = (custom_date - timedelta(days=1)).strftime('%Y-%m-%d') + ' 09:00:00'
        df['发布时间'] = pd.to_datetime(df['发布时间'])
        filtered_news = df[df['发布时间'] >= start_time]
        description = ""
        for index, row in filtered_news.iterrows():
            description += f"{row['发布时间'].strftime('%Y-%m-%d %H:%M:%S')} - {row['内容']}\n"
        return description
    except Exception as e:
        return f"获取新闻数据失败: {e}"

def create_folder_and_doc_path(custom_date):
    base_path = "C:/Users/jacky/Desktop/期货日报"
    folder_path = os.path.join(base_path, f"期货日报_{custom_date}")
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"文件夹 '{folder_path}' 已创建。")
    else:
        print(f"文件夹 '{folder_path}' 已存在。")
    base_filename = "恒力期货日报"
    filename = f"{base_filename}_{custom_date}.docx"
    doc_path = os.path.join(folder_path, filename)
    return doc_path, folder_path

def set_doc_style(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "楷体")
    normal.font.size = Pt(12)

def set_font_kaiti(paragraph):
    run = paragraph.add_run()
    run.font.name = '楷体'
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '楷体')
    run.font.size = Pt(12)
    return run

def create_report(custom_date_str, symbol, user_description, main_view):
    custom_date = datetime.strptime(custom_date_str, '%Y-%m-%d')
    doc_path, folder_path = create_folder_and_doc_path(custom_date_str)
    day_description, night_description, market_data = get_market_trend_data(symbol, custom_date)
    news_description = get_news_data(symbol, custom_date)

    if market_data.empty:
        print("无法生成报告，因为市场数据为空。")
        return None

    k_line_chart_path = create_k_line_chart(market_data, symbol, folder_path)

    doc = Document()
    set_doc_style(doc)

    title = doc.add_paragraph(f"{symbol}恒力期货日报{custom_date_str}")
    title_run = title.runs[0]
    title_run.font.size = Pt(14)
    title_run.bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    main_view_paragraph = doc.add_paragraph()
    main_view_run = main_view_paragraph.add_run("主要观点：")
    main_view_run.bold = True
    main_view_paragraph.add_run("\n" + main_view)

    core_logic = doc.add_paragraph()
    core_logic_run = core_logic.add_run("核心逻辑：")
    core_logic_run.bold = True
    core_logic.add_run("\n" + user_description)

    market_trend_paragraph = doc.add_paragraph()
    market_trend_run = market_trend_paragraph.add_run("前日走势：")
    market_trend_run.bold = True
    if k_line_chart_path:
        doc.add_picture(k_line_chart_path, width=Inches(6))
    market_trend_paragraph.add_run("\n" + day_description + "\n" + night_description)
    set_font_kaiti(market_trend_paragraph)

    news_paragraph = doc.add_paragraph()
    news_run = news_paragraph.add_run("今日新闻资讯：")
    news_run.bold = True
    set_font_kaiti(news_paragraph)
    news_paragraph.add_run("\n" + news_description)

    appendix = doc.add_paragraph()
    appendix_run = appendix.add_run("附录")
    appendix_run.bold = True

    doc.save(doc_path)
    return doc_path

st.title("期货日报生成")
st.write("created by 恒力期货上海分公司")

custom_date = st.date_input("请选择日期")
symbol = st.selectbox("请选择品种", ['铜', '铝', '铅', '锌', '镍', '锡'])
full_contract = st.text_input("请输入品种合约（如：cu2408）")

if st.button("生成K线图"):
    market_data = get_market_trend_data(full_contract, custom_date)[2]
    if not market_data.empty:
        k_line_chart_path = create_k_line_chart(market_data, full_contract, ".")
        st.image(k_line_chart_path)
        st.success("K线图生成成功")
    else:
        st.error("无法生成K线图，因为市场数据为空。")

user_description = st.text_area("请输入行情描述")
main_view = st.text_area("请输入主要观点")

if st.button("生成日报"):
    custom_date_str = custom_date.strftime('%Y-%m-%d')
    doc_path = create_report(custom_date_str, symbol, user_description, main_view)
    if doc_path:
        with open(doc_path, "rb") as f:
            st.download_button(
                label="下载日报",
                data=f,
                file_name=os.path.basename(doc_path),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.error("无法生成报告，因为市场数据为空。")
