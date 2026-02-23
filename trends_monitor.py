import os
import pandas as pd
from datetime import datetime, timedelta
import schedule
import time
import random
from querytrends import batch_get_queries, save_related_queries, RequestLimiter
import logging
import backoff
import argparse
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
            # 将高趋势分批处理，每批最多10个趋势
            batch_size = 10
            for i in range(0, len(high_rising_trends), batch_size):
                batch_trends = high_rising_trends[i:i + batch_size]
                batch_number = i // batch_size + 1
                total_batches = (len(high_rising_trends) + batch_size - 1) // batch_size
                
                alert_body = f"""
                <h2>📊 High Rising Trends Alert</h2>
                <hr>
                <h3>📌 Query Parameters:</h3>
                <ul>
                    <li>🕒 Time Range: {TRENDS_CONFIG['timeframe']}</li>
                    <li>🌍 Region: {TRENDS_CONFIG['geo'] or 'Global'}</li>
                </ul>
                <h3>📈 Significant Growth Trends:</h3>
                <table border="1" cellpadding="5" style="border-collapse: collapse;">
                    <tr>
                        <th>🔍 Base Keyword</th>
                        <th>🔗 Related Query</th>
                        <th>📈 Growth</th>
                    </tr>
                """
                
                for keyword, related_keywords, value in batch_trends:
                    alert_body += f"""
                    <tr>
                        <td><strong>🎯 {keyword}</strong></td>
                        <td>➡️ {related_keywords}</td>
                        <td align="right" style="color: #28a745;">⬆️ {value}%</td>
                    </tr>
                    """
                
                alert_body += "</table>"
                
                if batch_number < total_batches:
                    alert_body += f"<p><i>This is batch {batch_number} of {total_batches}. More results will follow.</i></p>"
                
                if not notification_manager.send_notification(
                    subject=f"📊 Rising Trends Alert ({batch_number}/{total_batches})",
                    body=alert_body
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
