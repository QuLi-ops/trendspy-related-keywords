import os
import pandas as pd
from datetime import datetime, timedelta
import schedule
import time
import random
from querytrends import batch_get_queries, get_interest_over_time, save_related_queries, RequestLimiter
import logging
import backoff
import argparse
import hashlib
import re
from html import escape
from config import (
    EMAIL_CONFIG, 
    KEYWORDS, 
    RATE_LIMIT_CONFIG, 
    SCHEDULE_CONFIG,
    MONITOR_CONFIG,
    LOGGING_CONFIG,
    STORAGE_CONFIG,
    TRENDS_CONFIG
)
from notification import NotificationManager

# Configure logging
logging.basicConfig(
    level=getattr(logging, LOGGING_CONFIG['level']),
    format=LOGGING_CONFIG['format'],
    handlers=[
        logging.FileHandler(LOGGING_CONFIG['log_file']),
        logging.StreamHandler()
    ]
)

# 创建请求限制器实例
request_limiter = RequestLimiter()

# 创建通知管理器实例
notification_manager = NotificationManager()

def create_daily_directory():
    """Create a directory for today's data"""
    today = datetime.now().strftime('%Y%m%d')
    directory = f"{STORAGE_CONFIG['data_dir_prefix']}{today}"
    if not os.path.exists(directory):
        os.makedirs(directory)
    return directory

def check_rising_trends(data, keyword, threshold=MONITOR_CONFIG['rising_threshold']):
    """Check if any rising trends exceed the threshold"""
    if not data or 'rising' not in data or data['rising'] is None:
        return []
    
    rising_trends = []
    df = data['rising']
    if isinstance(df, pd.DataFrame):
        for _, row in df.iterrows():
            if row['value'] > threshold:
                rising_trends.append((row['query'], row['value']))
    return rising_trends

def _safe_filename(value):
    """Create a filesystem-safe fragment for chart filenames."""
    cleaned = re.sub(r'[^a-zA-Z0-9._-]+', '_', value.strip())[:60]
    return cleaned or hashlib.md5(value.encode('utf-8')).hexdigest()[:12]

def _extract_interest_series(df, keyword):
    """Return the numeric interest series from a trendspy interest_over_time result."""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return None

    if keyword in df.columns:
        series = df[keyword]
    else:
        numeric_columns = [
            col for col in df.columns
            if col != 'isPartial' and pd.api.types.is_numeric_dtype(df[col])
        ]
        if not numeric_columns:
            return None
        series = df[numeric_columns[0]]

    series = pd.to_numeric(series, errors='coerce').fillna(0)
    if series.empty:
        return None
    return series

def summarize_trend_shape(series):
    """Create a short human-readable label for the 7-day curve shape."""
    if series is None or series.empty:
        return "无可用趋势数据"

    max_value = float(series.max())
    latest = float(series.iloc[-1])
    average = float(series.mean())
    active_points = int((series > 0).sum())
    peak_position = int(series.reset_index(drop=True).idxmax())
    last_position = len(series) - 1

    if max_value <= 0 or active_points == 0:
        return "7天内几乎没有搜索量"

    if peak_position < last_position * 0.75 and latest <= max_value * 0.35:
        return f"峰值已回落，当前 {latest:.0f}/100"

    if latest >= max_value * 0.75 and latest >= average * 1.5:
        return f"仍在高位，当前 {latest:.0f}/100"

    if active_points <= max(2, len(series) * 0.2):
        return f"短促尖峰，峰值 {max_value:.0f}/100"

    return f"相对平稳，均值 {average:.0f}/100"

def render_trend_chart(keyword, series, output_path):
    """Render a compact Google-Trends-style 7-day line chart."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.dates as mdates
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(5.4, 2.2), dpi=120)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    x_values = series.index
    y_values = series.values
    ax.plot(x_values, y_values, color='#4285f4', linewidth=2.4)
    ax.fill_between(x_values, y_values, 0, color='#4285f4', alpha=0.08)

    ax.set_ylim(0, 105)
    ax.set_yticks([0, 25, 50, 75, 100])
    ax.grid(axis='y', color='#e6e8eb', linewidth=0.8)
    ax.grid(axis='x', visible=False)
    ax.tick_params(axis='both', labelsize=8, colors='#6b7280', length=0)
    ax.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=3, maxticks=4))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))

    for spine in ax.spines.values():
        spine.set_visible(False)

    ax.set_title(keyword, loc='left', fontsize=11, color='#111827', pad=8)
    fig.tight_layout(pad=1.0)
    fig.savefig(output_path, format='png', bbox_inches='tight')
    plt.close(fig)

def build_trend_chart(keyword, directory):
    """Fetch 7-day trend data and render a chart image for one related query."""
    timeframe = MONITOR_CONFIG.get('chart_timeframe', 'now 7-d')
    try:
        df = get_interest_over_time(
            keyword,
            geo=TRENDS_CONFIG['geo'],
            timeframe=timeframe
        )
        series = _extract_interest_series(df, keyword)
        if series is None:
            return None, "无可用趋势数据"

        chart_filename = (
            f"trend_chart_{_safe_filename(keyword)}_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        )
        chart_path = os.path.join(directory, chart_filename)
        render_trend_chart(keyword, series, chart_path)
        return chart_path, summarize_trend_shape(series)
    except Exception as e:
        logging.warning(f"Failed to build trend chart for '{keyword}': {str(e)}")
        return None, "趋势图生成失败"

def build_alert_email(batch_trends, directory):
    """Build a mobile-friendly alert email with one 7-day chart per related query."""
    inline_images = {}
    timeframe = MONITOR_CONFIG.get('chart_timeframe', 'now 7-d')

    body = f"""
    <div style="font-family: Arial, Helvetica, sans-serif; color: #111827; line-height: 1.45;">
        <h2 style="margin: 0 0 12px 0;">High Rising Trends Alert</h2>
        <div style="font-size: 14px; color: #4b5563; margin-bottom: 18px;">
            Time Range: {escape(TRENDS_CONFIG['timeframe'])}<br>
            Region: {escape(TRENDS_CONFIG['geo'] or 'Global')}<br>
            Chart Range: {escape(timeframe)}
        </div>
    """

    for index, (keyword, related_keywords, value) in enumerate(batch_trends, start=1):
        chart_path, trend_summary = build_trend_chart(related_keywords, directory)
        cid = f"trend_chart_{hashlib.md5(f'{related_keywords}-{index}'.encode('utf-8')).hexdigest()}"

        if chart_path:
            inline_images[cid] = chart_path
            chart_html = (
                f'<img src="cid:{cid}" alt="{escape(related_keywords)} 7-day trend" '
                'style="display: block; width: 100%; max-width: 560px; height: auto; '
                'border: 1px solid #e5e7eb; border-radius: 6px;">'
            )
        else:
            chart_html = (
                '<div style="padding: 18px; border: 1px solid #e5e7eb; '
                'border-radius: 6px; color: #6b7280;">No chart available</div>'
            )

        body += f"""
        <div style="border: 1px solid #d9dde3; border-radius: 8px; padding: 14px; margin: 0 0 14px 0;">
            <div style="font-size: 18px; font-weight: 700; margin-bottom: 6px;">
                {escape(related_keywords)}
            </div>
            <div style="font-size: 14px; color: #4b5563; margin-bottom: 10px;">
                Base: {escape(keyword)} · Growth: <span style="color: #16a34a; font-weight: 700;">{value}%</span>
            </div>
            <div style="font-size: 14px; color: #111827; margin-bottom: 10px;">
                {escape(trend_summary)}
            </div>
            {chart_html}
        </div>
        """

    body += "</div>"
    return body, inline_images

def generate_daily_report(results, directory):
    """Generate a daily report in CSV format"""
    report_data = []
    
    for keyword, data in results.items():
        if data and isinstance(data.get('rising'), pd.DataFrame):
            rising_df = data['rising']
            for _, row in rising_df.iterrows():
                report_data.append({
                    'keyword': keyword,
                    'related_keywords': row['query'],
                    'value': row['value'],
                    'type': 'rising'
                })
        
        if data and isinstance(data.get('top'), pd.DataFrame):
            top_df = data['top']
            for _, row in top_df.iterrows():
                report_data.append({
                    'keyword': keyword,
                    'related_keywords': row['query'],
                    'value': row['value'],
                    'type': 'top'
                })
    
    if report_data:
        df = pd.DataFrame(report_data)
        filename = f"{STORAGE_CONFIG['report_filename_prefix']}{datetime.now().strftime('%Y%m%d')}.csv"
        report_file = os.path.join(directory, filename)
        df.to_csv(report_file, index=False)
        return report_file
    return None

def get_date_range_timeframe(timeframe):
    """Convert special timeframe formats to date range format
    
    Args:
        timeframe (str): Timeframe string like 'last-2-d' or 'last-3-d'
        
    Returns:
        str: Date range format string like '2024-01-01 2024-01-31'
    """
    if not timeframe.startswith('last-'):
        return timeframe
        
    try:
        # 解析天数
        days = int(timeframe.split('-')[1])
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days)
        # 格式化日期字符串
        return f"{start_date.strftime('%Y-%m-%d')} {end_date.strftime('%Y-%m-%d')}"
    except (ValueError, IndexError):
        logging.warning(f"Invalid timeframe format: {timeframe}, falling back to 'now 1-d'")
        return 'now 1-d'

def process_keywords_batch(keywords_batch, directory, all_results, high_rising_trends, timeframe):
    """处理一批关键词"""
    try:
        logging.info(f"Processing batch of {len(keywords_batch)} keywords")
        logging.info(f"Query parameters: timeframe={timeframe}, geo={TRENDS_CONFIG['geo'] or 'Global'}")
        
        # 使用传入的 timeframe 参数
        results = get_trends_with_retry(keywords_batch, timeframe)
        
        for keyword, data in results.items():
            if data:
                filename = save_related_queries(keyword, data)
                if filename:
                    os.rename(filename, os.path.join(directory, filename))
                
                rising_trends = check_rising_trends(data, keyword)
                if rising_trends:
                    high_rising_trends.extend([(keyword, related_keywords, value) 
                                             for related_keywords, value in rising_trends])
                
                all_results[keyword] = data
        
        return True
    except Exception as e:
        logging.error(f"Error processing batch: {str(e)}")
        return False

@backoff.on_exception(
    backoff.expo,
    Exception,
    max_tries=RATE_LIMIT_CONFIG['max_retries'],
    jitter=backoff.full_jitter
)
def get_trends_with_retry(keywords_batch, timeframe):
    """使用重试机制获取趋势数据"""
    return batch_get_queries(
        keywords_batch,
        timeframe=timeframe,  # 使用传入的 timeframe
        geo=TRENDS_CONFIG['geo'],
        delay_between_queries=random.uniform(
            RATE_LIMIT_CONFIG['min_delay_between_queries'],
            RATE_LIMIT_CONFIG['max_delay_between_queries']
        )
    )

def process_trends():
    """Main function to process trends data"""
    try:
        logging.info("Starting daily trends processing")
        
        # 处理特殊的 timeframe 格式
        timeframe = TRENDS_CONFIG['timeframe']
        actual_timeframe = get_date_range_timeframe(timeframe)
        
        logging.info(f"Using configuration: timeframe={actual_timeframe}, geo={TRENDS_CONFIG['geo'] or 'Global'}")
        directory = create_daily_directory()
        
        all_results = {}
        high_rising_trends = []
        
        # 将关键词分批处理，使用实际的 timeframe
        for i in range(0, len(KEYWORDS), RATE_LIMIT_CONFIG['batch_size']):
            keywords_batch = KEYWORDS[i:i + RATE_LIMIT_CONFIG['batch_size']]
            # 传递实际的 timeframe 到查询函数
            success = process_keywords_batch(
                keywords_batch, 
                directory, 
                all_results, 
                high_rising_trends,
                actual_timeframe
            )
            
            if not success:
                logging.error(f"Failed to process batch starting with keyword: {keywords_batch[0]}")
                continue
            
            # 如果不是最后一批，等待一段时间再处理下一批
            if i + RATE_LIMIT_CONFIG['batch_size'] < len(KEYWORDS):
                wait_time = RATE_LIMIT_CONFIG['batch_interval'] + random.uniform(0, 60)
                logging.info(f"Waiting {wait_time:.1f} seconds before processing next batch...")
                time.sleep(wait_time)

        # Generate and send daily report
        report_file = generate_daily_report(all_results, directory)
        if report_file:
            report_body = """
            <h2>Daily Trends Report</h2>
            <p>Please find attached the daily trends report.</p>
            <p>Query Parameters:</p>
            <ul>
            <li>Time Range: {}</li>
            <li>Region: {}</li>
            </ul>
            <p>Summary:</p>
            <ul>
            <li>Total keywords processed: {}</li>
            <li>Successful queries: {}</li>
            <li>Failed queries: {}</li>
            </ul>
            """.format(
                TRENDS_CONFIG['timeframe'],
                TRENDS_CONFIG['geo'] or 'Global',
                len(KEYWORDS),
                len(all_results),
                len(KEYWORDS) - len(all_results)
            )
            if not notification_manager.send_notification(
                subject=f"Daily Trends Report - {datetime.now().strftime('%Y-%m-%d')}",
                body=report_body,
                attachments=[report_file]
            ):
                logging.warning("Failed to send daily report, but data collection completed")
        
        # Send alerts for high rising trends
        if high_rising_trends:
            # 将高趋势分批处理，邮件内每个词都带 7 天趋势图
            batch_size = 10
            for i in range(0, len(high_rising_trends), batch_size):
                batch_trends = high_rising_trends[i:i + batch_size]
                batch_number = i // batch_size + 1
                total_batches = (len(high_rising_trends) + batch_size - 1) // batch_size

                alert_body, inline_images = build_alert_email(batch_trends, directory)
                
                if batch_number < total_batches:
                    alert_body += f"<p><i>This is batch {batch_number} of {total_batches}. More results will follow.</i></p>"
                
                if not notification_manager.send_notification(
                    subject=f"📊 Rising Trends Alert ({batch_number}/{total_batches})",
                    body=alert_body,
                    inline_images=inline_images
                ):
                    logging.warning(f"Failed to send alert notification for batch {batch_number}, but data collection completed")
                
                # 添加短暂延迟，避免消息发送过快
                time.sleep(2)
        
        logging.info("Daily trends processing completed successfully")
        return True
    except Exception as e:
        logging.error(f"Error in trends processing: {str(e)}")
        notification_manager.send_notification(
            subject="❌ Error in Trends Processing",
            body=f"<p>An error occurred during trends processing:</p><pre>{str(e)}</pre>"
        )
        return False

def run_scheduler():
    """Run the scheduler"""
    # 从配置中获取小时和分钟
    schedule_hour = SCHEDULE_CONFIG['hour']
    schedule_minute = SCHEDULE_CONFIG.get('minute', 0)  # 默认为0分钟
    
    # 添加随机延迟（如果配置了的话）
    if SCHEDULE_CONFIG.get('random_delay_minutes', 0) > 0:
        random_minutes = random.randint(0, SCHEDULE_CONFIG['random_delay_minutes'])
        schedule_minute = (schedule_minute + random_minutes) % 60
        # 如果分钟数超过59，需要调整小时数
        schedule_hour = (schedule_hour + (schedule_minute + random_minutes) // 60) % 24
    
    schedule_time = f"{schedule_hour:02d}:{schedule_minute:02d}"
    
    schedule.every().day.at(schedule_time).do(process_trends)
    
    logging.info(f"Scheduler started. Will run daily at {schedule_time}")
    
    # 如果启动时间接近计划执行时间，等待到下一天
    now = datetime.now()
    scheduled_time = now.replace(hour=schedule_hour, minute=schedule_minute, second=0, microsecond=0)
    
    if now >= scheduled_time:
        logging.info("Current time is past scheduled time, waiting for tomorrow")
        next_run = scheduled_time + timedelta(days=1)
        time.sleep((next_run - now).total_seconds())
    
    while True:
        schedule.run_pending()
        time.sleep(60)

if __name__ == "__main__":
    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(description='Google Trends Monitor')
    parser.add_argument('--test', action='store_true', 
                      help='立即运行一次数据收集，而不是等待计划时间')
    parser.add_argument('--keywords', nargs='+',
                      help='测试时要查询的关键词列表，如果不指定则使用配置文件中的关键词')
    args = parser.parse_args()

    # 检查邮件配置
    if not all([
        EMAIL_CONFIG['sender_email'],
        EMAIL_CONFIG['sender_password'],
        EMAIL_CONFIG['recipient_email']
    ]):
        logging.error("Please configure email settings in config.py before running")
        exit(1)
    
    # 如果是测试模式
    if args.test:
        logging.info("Running in test mode...")
        if args.keywords:
            # 临时替换配置文件中的关键词
            global KEYWORDS
            KEYWORDS = args.keywords
            logging.info(f"Using test keywords: {KEYWORDS}")
        process_trends()
    else:
        # 正常的计划任务模式
        run_scheduler() 
