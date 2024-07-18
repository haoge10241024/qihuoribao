import os
import akshare as ak
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.pyplot as plt
from matplotlib import rcParams
import mplfinance as mpf
import streamlit as st

# 使用系统字体 SimHei
rcParams['font.sans-serif'] = ['SimHei']
rcParams['axes.unicode_minus'] = False

# 创建文件夹和文档保存路径
def create_folder_and_doc_path(custom_date):
    base_path = "C:/Users/jacky/Desktop/期货日报"
    folder_path = os.path.join(base_path, f"期货日报_{custom_date}")
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    base_filename = "恒力期货日报"
    filename = f"{base_filename}_{custom_date}.docx"
    doc_path = os.path.join(folder_path, filename)
    return doc_path, folder_path

# 设置文档样式
def set_doc_style(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "楷体")
    normal.font.size = Pt(12)

# 获取当天行情概述数据
def get_market_trend_data(symbol, custom_date):
    try:
        today = custom_date
        yesterday = today - timedelta(days=1)
        start_time = yesterday.strftime('%Y-%m-%d') + ' 21:00:00'
        end_time = today.strftime('%Y-%m-%d') + ' 23:00:00'
        df = ak.futures_zh_minute_sina(symbol=symbol, period="1")
        df['datetime'] = pd.to_datetime(df['datetime'])
        filtered_data = df[(df['datetime'] >= start_time) & (df['datetime'] <= end_time)]
        return filtered_data
    except Exception as e:
        print(f"获取市场走势数据失败: {e}")
        return pd.DataFrame()

# 创建K线图
def create_k_line_chart(data, symbol, folder_path):
    if data.empty:
        print("数据为空，无法生成K线图。")
        return None
    data.set_index('datetime', inplace=True)
    data = data[['open', 'high', 'low', 'close']]
    data.columns = ['Open', 'High', 'Low', 'Close']
    fig, ax = plt.subplots(figsize=(10, 6))
    mpf.plot(data, type='candle', style='charles', ax=ax)
    ax.set_title(f'{symbol} 当日K线图')
    ax.set_ylabel('价格 (元/吨)')
    k_line_chart_path = os.path.join(folder_path, 'k_line_chart.png')
    plt.savefig(k_line_chart_path)
    plt.close(fig)
    return k_line_chart_path

# 获取新闻数据
def get_news_data():
    try:
        df = ak.futures_news_shmet(symbol="铜")
        latest_news = df.tail(30)
        description = ""
        for index, row in latest_news.iterrows():
            description += f"{row['发布时间'].strftime('%Y-%m-%d %H:%M:%S')} - {row['内容']}\n"
        return description
    except Exception as e:
        return f"获取新闻数据失败: {e}"

# 设置楷体字体
def set_font_kaiti(paragraph):
    run = paragraph.add_run()
    run.font.name = '楷体'
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '楷体')
    run.font.size = Pt(12)
    return run

# 创建报告
def create_report(custom_date_str, symbol, user_description, main_view):
    custom_date = datetime.strptime(custom_date_str, '%Y-%m-%d')
    doc_path, folder_path = create_folder_and_doc_path(custom_date_str)
    market_trend_description, night_trend_description, market_data = get_market_trend_data(symbol=symbol, custom_date=custom_date)
    news_description = get_news_data()
    
    roll_yield_chart_path = os.path.join(folder_path, 'roll_yield_chart.png')
    fetch_and_plot_futures_data(symbol, roll_yield_chart_path)
    
    k_line_chart_path = create_k_line_chart(market_data, symbol, folder_path)

    doc = Document()
    set_doc_style(doc)

    # 添加标题
    title = doc.add_paragraph("恒力期货日报")
    title_run = title.runs[0]
    title_run.font.size = Pt(14)
    title_run.bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 添加主要观点段落
    main_view_paragraph = doc.add_paragraph()
    main_view_run = main_view_paragraph.add_run("主要观点：")
    main_view_run.bold = True
    main_view_paragraph.add_run("\n" + main_view)

    # 添加核心逻辑段落
    core_logic = doc.add_paragraph()
    core_logic_run = core_logic.add_run("核心逻辑：")
    core_logic_run.bold = True

    # 添加前日走势
    market_trend_paragraph = doc.add_paragraph()
    market_trend_run = market_trend_paragraph.add_run("前日走势：")
    market_trend_run.bold = True

    if k_line_chart_path:
        doc.add_picture(k_line_chart_path, width=Inches(6))
    market_trend_paragraph.add_run("\n" + user_description)
    set_font_kaiti(market_trend_paragraph)

    # 添加期限结构图
    if roll_yield_chart_path and os.path.exists(roll_yield_chart_path):
        doc.add_paragraph("期限结构图")
        doc.add_picture(roll_yield_chart_path, width=Inches(6))

    # 添加今日新闻资讯
    news_paragraph = doc.add_paragraph()
    news_run = news_paragraph.add_run("今日新闻资讯：")
    news_run.bold = True
    set_font_kaiti(news_paragraph)
    news_paragraph.add_run("\n" + news_description)

    # 添加附录
    appendix = doc.add_paragraph()
    appendix_run = appendix.add_run("附录")
    appendix_run.bold = True

    # 保存文档
    doc.save(doc_path)
    return doc_path

# Streamlit应用
st.title("期货日报生成器")
st.write("请选择日期和品种，输入主要观点和行情描述，然后点击生成日报")

custom_date = st.date_input("请选择日期")
symbol = st.selectbox("请选择品种", ['cu', 'al', 'pb', 'zn', 'ni', 'sn'])

if st.button("生成K线图"):
    custom_date_str = custom_date.strftime('%Y-%m-%d')
    market_data = get_market_trend_data(symbol, custom_date)
    k_line_chart_path = create_k_line_chart(market_data, symbol, ".")

    if k_line_chart_path:
        st.image(k_line_chart_path, caption="前日K线图")
    st.write("前日走势：")
    st.write(market_data)

user_description = st.text_area("请输入行情描述（自动生成或自行编辑）")
main_view = st.text_area("请输入主要观点")

if st.button("生成日报"):
    custom_date_str = custom_date.strftime('%Y-%m-%d')
    doc_path = create_report(custom_date_str, symbol, user_description, main_view)
    with open(doc_path, "rb") as f:
        st.download_button(
            label="下载日报",
            data=f,
            file_name=os.path.basename(doc_path),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
