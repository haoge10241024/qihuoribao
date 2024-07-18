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

# 设置字体为DejaVu Sans
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# 创建文件夹和文档保存路径
def create_folder_and_doc_path(custom_date, symbol):
    base_path = "C:/Users/jacky/Desktop/期货日报"
    folder_path = os.path.join(base_path, f"期货日报_{custom_date}")
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    base_filename = f"{symbol}恒力期货日报_{custom_date}"
    filename = f"{base_filename}.docx"
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
        
        if filtered_data.empty:
            print("市场数据为空")
            return "", "", pd.DataFrame()

        # 获取开盘价和收盘价
        day_open_price = filtered_data.iloc[0]['open']  # 前一晚21:00开盘价
        day_close_price = filtered_data[filtered_data['datetime'] <= today.strftime('%Y-%m-%d') + ' 15:00:00'].iloc[-1]['close']  # 15:00收盘价

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

        # 获取夜盘走势
        night_start_time = today.strftime('%Y-%m-%d') + ' 21:00:00'
        night_end_time = (today + timedelta(days=1)).strftime('%Y-%m-%d') + ' 01:00:00'
        night_data = df[(df['datetime'] >= night_start_time) & (df['datetime'] <= night_end_time)]
        if night_data.empty:
            night_description = "夜盘数据不可用。"
        else:
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
        return f"获取市场走势数据失败: {e}", "", pd.DataFrame()

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
    ax.set_ylabel('价格 (元/吨)')
    k_line_chart_path = os.path.join(folder_path, 'k_line_chart.png')
    plt.savefig(k_line_chart_path)
    plt.close(fig)
    return k_line_chart_path

# 获取新闻数据
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

# 设置楷体字体
def set_font_kaiti(paragraph):
    run = paragraph.add_run()
    run.font.name = '楷体'
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '楷体')
    run.font.size = Pt(12)
    return run

# 获取合约期限结构数据并生成图表
def fetch_and_plot_futures_data(commodity_name, output_filename):
    continuous_contracts = [
        "V0", "P0", "B0", "M0", "I0", "JD0", "L0", "PP0", "FB0", "BB0", "Y0",
        "C0", "A0", "J0", "JM0", "CS0", "EG0", "RR0", "EB0", "PG0", "LH0",
        "TA0", "OI0", "RS0", "RM0", "WH0", "JR0", "SR0", "CF0", "RI0", "MA0",
        "FG0", "LR0", "SF0", "SM0", "CY0", "AP0", "CJ0", "UR0", "SA0", "PF0",
        "PK0", "SH0", "PX0", "FU0", "SC0", "AL0", "RU0", "ZN0", "CU0", "AU0",
        "RB0", "WR0", "PB0", "AG0", "BU0", "HC0", "SN0", "NI0", "SP0", "NR0",
        "SS0", "LU0", "BC0", "AO0", "BR0", "EC0", "IF0", "TF0", "IH0", "IC0",
        "TS0", "IM0", "SI0", "LC0"
    ]

    all_symbols = []

    try:
        futures_zh_realtime_df = ak.futures_zh_realtime(symbol=commodity_name)
        symbols = futures_zh_realtime_df['symbol'].tolist()
        all_symbols.extend([s for s in symbols if s not in continuous_contracts])
    except Exception as e:
        print(f"Error fetching realtime data for {commodity_name}: {e}")
        return

    all_symbols = sorted(set(all_symbols))

    symbol_close_prices = {}

    end_date = datetime.now()
    start_date = end_date - timedelta(days=9)  # 获取数据的天数，改为9天

    for symbol in all_symbols:
        try:
            futures_zh_daily_sina_df = ak.futures_zh_daily_sina(symbol=symbol)
            futures_zh_daily_sina_df['date'] = pd.to_datetime(futures_zh_daily_sina_df['date'])
            recent_days_df = futures_zh_daily_sina_df[(futures_zh_daily_sina_df['date'] >= start_date) & (futures_zh_daily_sina_df['date'] <= end_date)]
            symbol_close_prices[symbol] = recent_days_df.set_index('date')['close']
        except Exception as e:
            print(f"Error fetching daily data for {symbol}: {e}")

    all_data = pd.DataFrame(symbol_close_prices)

    dates = all_data.index.unique()
    num_dates = len(dates)
    num_cols = 3
    num_rows = (num_dates + num_cols - 1) // num_cols

    fig, axes = plt.subplots(num_rows, num_cols, figsize=(18, 6 * num_rows), sharex=False, sharey=False)
    axes = axes.flatten()

    for i, current_date in enumerate(dates):
        ax = axes[i]
        date_str = current_date.strftime('%Y-%m-%d')
        prices_on_date = all_data.loc[current_date]
        
        ax.plot(prices_on_date.index, prices_on_date.values, marker='o')
        ax.set_title(date_str)
        ax.set_xticks(range(len(prices_on_date.index)))
        ax.set_xticklabels(prices_on_date.index, rotation=45)
        ax.set_ylabel('')
        ax.set_xlabel('')

    for j in range(i + 1, len(axes)):
        fig.delaxes(axes[j])

    plt.tight_layout()
    plt.savefig(output_filename, dpi=300)
    plt.show()

# 创建报告
def create_report(custom_date_str, symbol, user_description, main_view):
    custom_date = datetime.strptime(custom_date_str, '%Y-%m-%d')
    doc_path, folder_path = create_folder_and_doc_path(custom_date_str, symbol)
    market_trend_description, night_trend_description, market_data = get_market_trend_data(symbol=symbol, custom_date=custom_date)
    
    if market_data.empty:
        st.error("无法生成报告，因为市场数据为空。")
        return None
    
    news_description = get_news_data(symbol, custom_date)
    
    roll_yield_chart_path = os.path.join(folder_path, 'roll_yield_chart.png')
    fetch_and_plot_futures_data(symbol, roll_yield_chart_path)
    
    k_line_chart_path = create_k_line_chart(market_data, symbol, folder_path)

    doc = Document()
    set_doc_style(doc)

    # 添加标题
    title = doc.add_paragraph(f"{symbol}恒力期货日报 {custom_date_str}")
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
st.title("期货日报生成")
st.write("created by 恒力期货上海分公司")

custom_date = st.date_input("请选择日期")
symbol = st.selectbox("请选择品种", ['铜', '铝', '铅', '锌', '镍', '锡'])
full_contract = st.text_input("请输入完整品种合约（如：CU2408）")

if st.button("生成K线图"):
    custom_date_str = custom_date.strftime('%Y-%m-%d')
    day_description, night_description, market_data = get_market_trend_data(full_contract, custom_date)
    k_line_chart_path = create_k_line_chart(market_data, full_contract, ".")

    if k_line_chart_path:
        st.image(k_line_chart_path, caption="前日K线图")
    else:
        st.error("无法生成K线图，因为市场数据为空。")
    
    st.write("前日走势：")
    st.write(day_description)
    st.write(night_description)

user_description = st.text_area("请输入行情描述（自动生成或自行编辑）")
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
