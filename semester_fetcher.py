import requests
from bs4 import BeautifulSoup
import json
import re
from urllib.parse import unquote
from datetime import datetime, timedelta, timezone

def parse_date_timestamp(date_str):
    """
    解析 .NET Date 格式的时间戳
    格式: /Date(1756656000000+0800)/
    返回: datetime 对象（带时区信息）
    """
    # 正则解析时间戳与时区偏移
    pattern = r"/?Date\((\d+)([+-]\d{4})?\)/?"
    match = re.search(pattern, date_str)

    if not match:
        raise ValueError(f"Invalid .NET date format: {date_str}")

    millis = int(match.group(1))   # 毫秒时间戳
    tz_str = match.group(2)        # 时区，例如 +0800

    # 处理时区
    if tz_str:
        sign = 1 if tz_str[0] == "+" else -1
        hours = int(tz_str[1:3])
        minutes = int(tz_str[3:5])
        tz = timezone(sign * timedelta(hours=hours, minutes=minutes))
    else:
        tz = timezone.utc

    # 生成 datetime
    timestamp_sec = millis / 1000
    dt = datetime.fromtimestamp(timestamp_sec, tz)

    return dt

def fetch_semester_info():
    """
    从 https://bksy.bjtu.edu.cn/Admin/SemesterTranPage.aspx 获取学期信息
    返回: (semester_start, rest_weeks)
    - semester_start: datetime 对象，表示学期第一周周一的日期
    - rest_weeks: 列表，格式为 [(after_week, rest_count), ...]
    """
    url = "https://bksy.bjtu.edu.cn/Admin/SemesterTranPage.aspx"
    
    try:
        print("正在从网页获取学期信息...")
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        response.encoding = 'utf-8'
        
        # 解析 HTML
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 查找 id 为 hidJson 的 input 元素
        input_element = soup.find('input', {'id': 'hidJson'})
        if not input_element:
            raise ValueError("未找到 id 为 hidJson 的 input 元素")
        
        # 获取 value 值
        encoded_value = input_element.get('value', '')
        if not encoded_value:
            raise ValueError("hidJson input 元素的 value 为空")
        
        # URL 解码
        decoded_value = unquote(encoded_value)
        
        # 解析 JSON
        try:
            semester_data = json.loads(decoded_value)
        except json.JSONDecodeError as e:
            raise ValueError(f"JSON 解析失败: {e}")
        
        # 提取学期开始日期和休息周信息
        semester_start, rest_weeks = extract_semester_info(semester_data)
        
        print(f"学期开始日期: {semester_start.strftime('%Y年%m月%d日')}")
        if rest_weeks:
            rest_info = ", ".join([f"第{week}周后休息{count}周" for week, count in rest_weeks])
            print(f"休息周信息: {rest_info}")
        else:
            print("无休息周")
        
        return semester_start, rest_weeks
        
    except requests.RequestException as e:
        raise Exception(f"网络请求失败: {e}")
    except Exception as e:
        raise Exception(f"获取学期信息失败: {e}")

def extract_semester_info(semester_data):
    """
    从学期数据中提取学期开始日期和休息周信息
    :param semester_data: JSON 数据，格式如 decoded.json
    :return: (semester_start, rest_weeks)
    """
    semester_start = None
    rest_weeks = []
    
    # semester_data 是一个列表，包含多个学期对象
    # 每个对象有 Id 和 Json 字段
    # Json 是一个列表，包含每周的信息
    
    # 遍历所有学期数据
    for semester_obj in semester_data:
        if 'Json' not in semester_obj:
            continue
        
        weeks_data = semester_obj['Json']
        
        # 找到第一个有 SemesterName 且 Week 为 "第1教学周" 的项
        first_teaching_week = None
        for week_info in weeks_data:
            week_str = week_info.get('Week', '')
            semester_name = week_info.get('SemesterName', '')
            
            # 查找第1教学周
            if week_str == '第1教学周' and semester_name:
                dt_str = week_info.get('DT', '')
                if dt_str:
                    try:
                        week_date = parse_date_timestamp(dt_str)
                        semester_start = week_date
                        # # 找到 week_date 所在周的周一
                        # current_weekday = week_date.weekday()  # 0=Monday, 6=Sunday
                        # semester_start = week_date - timedelta(days=current_weekday)
                        
                        first_teaching_week = week_info
                        break
                    except Exception as e:
                        print(f"警告: 解析日期失败: {e}")
                        continue
        
        # 如果找到了学期开始日期，继续处理这个学期的休息周信息
        if semester_start and first_teaching_week:
            # 先提取所有周的周数信息，得到一个只有数字和"休"的列表
            week_numbers = []
            first_teaching_week_index = None
            current_semester_name = first_teaching_week.get('SemesterName', '')
            
            for idx, week_info in enumerate(weeks_data):
                # 如果遇到新的学期（SemesterName 变化且不为空），停止处理
                semester_name = week_info.get('SemesterName', '')
                if semester_name and semester_name != current_semester_name:
                    break
                
                week_str = week_info.get('Week', '').strip()
                if week_str:
                    # 提取周数：strip('第').strip('教学周')
                    week_num = week_str.strip('第').strip('教学周').strip()
                    week_numbers.append(week_num)
                    
                    # 记录第1教学周的索引
                    if week_info == first_teaching_week:
                        first_teaching_week_index = len(week_numbers) - 1
            
            # 从第1教学周开始处理休息周信息
            if first_teaching_week_index is not None:
                i = first_teaching_week_index
                while i < len(week_numbers):
                    if week_numbers[i] == '休':
                        # 查找前面最近的一个教学周（数字）
                        after_week = None
                        for j in range(i - 1, -1, -1):
                            if week_numbers[j].isdigit():
                                after_week = int(week_numbers[j])
                                break
                        
                        if after_week is not None:
                            # 计算休息周的数量（可能连续多周）
                            rest_count = 1
                            i += 1
                            while i < len(week_numbers) and week_numbers[i] == '休':
                                rest_count += 1
                                i += 1
                            
                            rest_weeks.append((after_week, rest_count))
                        else:
                            i += 1
                    else:
                        i += 1
            
            # 如果找到了学期开始日期，就退出循环
            if semester_start:
                break
    
    if not semester_start:
        raise ValueError("未找到学期开始日期（第1教学周）")
    
    # 去重并排序休息周信息
    rest_weeks = list(set(rest_weeks))
    rest_weeks.sort(key=lambda x: x[0])
    
    return semester_start, rest_weeks

if __name__ == "__main__":
    try:
        semester_start, rest_weeks = fetch_semester_info()
        print(f"\n学期开始日期: {semester_start.strftime('%Y-%m-%d')}")
        print(f"休息周信息: {rest_weeks}")
    except Exception as e:
        print(f"错误: {e}")

