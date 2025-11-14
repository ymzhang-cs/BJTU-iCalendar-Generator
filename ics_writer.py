from ics import Calendar, Event
from ics.grammar.parse import ContentLine
from datetime import datetime, timedelta
import pytz
import tkinter as tk
from tkinter import filedialog
import ctypes
import os

# 添加时区 Asia/Shanghai
SHANGHAI_TZ = pytz.timezone("Asia/Shanghai")

# 课程时间段对应的具体时间
TIME_SLOTS = {
    1: "08:00",
    2: "10:10",
    3: "12:10",
    4: "14:10",
    5: "16:20",
    6: "19:00",
    7: "21:00",
}

# 错峰上课时间段
# 思源西楼、逸夫楼第二节课错后20分钟，其余不变
STAGGERED_KEYWORD = ['思源西楼','逸夫教学楼']
STAGGERED_TIME_SLOTS = {
    1: "08:00",
    2: "10:30",
    3: "12:10",
    4: "14:10",
    5: "16:20",
    6: "19:00",
    7: "21:00",
}

# 星期映射 (iCalendar 格式)
WEEKDAY_MAP = {
    1: "MO",
    2: "TU",
    3: "WE",
    4: "TH",
    5: "FR",
    6: "SA",
    7: "SU"
}

class Writer:
    def __init__(self, data, semester_start, rest_weeks=None):
        """
        :param data: 课程数据列表
        :param semester_start: 学期开始日期 (datetime 类型)
        :param rest_weeks: 休息周信息列表，格式为 [(after_week, rest_count), ...]
                          例如 [(3, 1)] 表示第3周后休息1周
        """
        self.data = data
        self.semester_start = semester_start  # 例如 datetime(2025, 3, 3)
        self.rest_weeks = rest_weeks if rest_weeks else []
        
        # 构建逻辑周次到实际周次的映射表
        # 映射表格式：{逻辑周次: 实际周次}
        # 同时记录哪些实际周次是休息周
        self.logical_to_actual = {}
        self.rest_actual_weeks = set()  # 记录哪些实际周次是休息周
        self._build_week_mapping()
    
    def _build_week_mapping(self):
        """构建逻辑周次到实际周次的映射表"""
        if not self.rest_weeks:
            # 没有休息周，逻辑周次和实际周次一一对应
            return
        
        # 计算最大逻辑周次（假设最大为100，实际应该根据课程数据动态计算）
        max_logical_week = 100
        
        actual_week = 1
        for logical_week in range(1, max_logical_week + 1):
            # 映射逻辑周次到实际周次
            self.logical_to_actual[logical_week] = actual_week
            actual_week += 1
            
            # 检查是否在某个休息周之后
            for after_week, rest_count in self.rest_weeks:
                if logical_week == after_week:
                    # 这是休息周前的最后一周，下一周是休息周
                    for rest_idx in range(rest_count):
                        rest_actual_week = actual_week
                        self.rest_actual_weeks.add(rest_actual_week)
                        # 休息周也映射到 after_week（用于生成课表时使用相同的逻辑周次）
                        actual_week += 1
    
    def logical_to_actual_week(self, logical_week):
        """将逻辑周次转换为实际周次"""
        if not self.rest_weeks:
            return logical_week
        return self.logical_to_actual.get(logical_week, logical_week)
    
    def get_all_actual_weeks_for_course(self, weeks_data):
        """获取课程对应的所有实际周次（包括休息周）"""
        actual_weeks = []
        
        if weeks_data["type"] == "continuous":
            start = weeks_data["data"]["start"]
            end = weeks_data["data"]["end"]
            for logical_week in range(start, end + 1):
                actual_week = self.logical_to_actual_week(logical_week)
                actual_weeks.append(actual_week)
                # 检查是否有休息周需要添加（休息周紧跟在 after_week 之后）
                for after_week, rest_count in self.rest_weeks:
                    if logical_week == after_week:
                        # 在休息周范围内，添加休息周
                        # 休息周的实际周次 = after_week 的实际周次 + 1, +2, ...
                        base_actual_week = self.logical_to_actual_week(after_week)
                        for i in range(rest_count):
                            rest_actual_week = base_actual_week + i + 1
                            if rest_actual_week not in actual_weeks:
                                actual_weeks.append(rest_actual_week)
        
        elif weeks_data["type"] == "discontinuous":
            logical_weeks = weeks_data["data"]
            for logical_week in logical_weeks:
                actual_week = self.logical_to_actual_week(logical_week)
                actual_weeks.append(actual_week)
                # 检查是否有休息周需要添加
                for after_week, rest_count in self.rest_weeks:
                    if logical_week == after_week:
                        base_actual_week = self.logical_to_actual_week(after_week)
                        for i in range(rest_count):
                            rest_actual_week = base_actual_week + i + 1
                            if rest_actual_week not in actual_weeks:
                                actual_weeks.append(rest_actual_week)
        
        elif weeks_data["type"] == "interval":
            start = weeks_data["data"]["start"]
            interval = weeks_data["data"]["interval"]
            count = weeks_data["data"]["count"]
            logical_weeks = [start + i * interval for i in range(count)]
            for logical_week in logical_weeks:
                actual_week = self.logical_to_actual_week(logical_week)
                actual_weeks.append(actual_week)
                # 检查是否有休息周需要添加
                for after_week, rest_count in self.rest_weeks:
                    if logical_week == after_week:
                        base_actual_week = self.logical_to_actual_week(after_week)
                        for i in range(rest_count):
                            rest_actual_week = base_actual_week + i + 1
                            if rest_actual_week not in actual_weeks:
                                actual_weeks.append(rest_actual_week)
        
        return sorted(actual_weeks)

    def generate_ics(self):
        """生成 ICS 日历"""
        cal = Calendar()

        for course in self.data:
            course_name = course["name"]
            teacher = course["teacher"]
            location = course["location"]
            weekday = course["time"]["weekday"]
            lesson = course["time"]["lesson"]
            weeks_data = course["weeks"]

            # 计算上课开始时间
            if any(keyword in location for keyword in STAGGERED_KEYWORD):
                start_time = STAGGERED_TIME_SLOTS.get(lesson)
            else:
                start_time = TIME_SLOTS.get(lesson)
            if not start_time:
                continue  # 避免无效时间段

            # 获取所有实际周次（包括休息周）
            actual_weeks = self.get_all_actual_weeks_for_course(weeks_data)
            if not actual_weeks:
                continue
            
            # 计算课程首次上课日期（使用第一个实际周次）
            first_actual_week = actual_weeks[0]
            first_class_date = self.semester_start + timedelta(days=(first_actual_week - 1) * 7 + (weekday - 1))
            start_dt = datetime.combine(first_class_date, datetime.strptime(start_time, "%H:%M").time())
            
            # 转换为 Asia/Shanghai 时区
            start_dt = SHANGHAI_TZ.localize(start_dt).astimezone(pytz.utc)
            end_dt = start_dt + timedelta(minutes=110 if lesson != 7 else 50)

            event = Event()
            event.name = f"{course_name} - {teacher}"
            event.begin = start_dt
            event.end = end_dt
            event.location = location

            # 生成 RRULE（基于实际周次）
            rrule_result = self.get_rrule_from_actual_weeks(actual_weeks, weekday, start_time)
            if rrule_result:
                rrule, exdates = rrule_result
                event.extra.append(ContentLine(name="RRULE", value=rrule))
                # 添加 EXDATE 排除不需要的周次
                for exdate in exdates:
                    event.extra.append(ContentLine(name="EXDATE", value=exdate))

            cal.events.add(event)

        return cal

    def get_first_week(self, weeks_data):
        """获取课程的第一次上课周"""
        if weeks_data["type"] == "continuous":
            return weeks_data["data"]["start"]
        elif weeks_data["type"] == "discontinuous":
            return min(weeks_data["data"])
        elif weeks_data["type"] == "interval":
            return weeks_data["data"]["start"]
        return 1  # 默认第 1 周

    def get_rrule(self, weeks_data, weekday):
        """生成 RRULE 规则（旧方法，保留用于兼容）"""
        week_day = WEEKDAY_MAP[weekday]

        if weeks_data["type"] == "continuous":
            start = weeks_data["data"]["start"]
            end = weeks_data["data"]["end"]
            count = end - start + 1
            return f"FREQ=WEEKLY;BYDAY={week_day};COUNT={count}"

        elif weeks_data["type"] == "discontinuous":
            weeks = weeks_data["data"]
            weeks_str = ",".join(str(w) for w in weeks)
            return f"FREQ=WEEKLY;BYDAY={week_day};BYSETPOS={weeks_str}"

        elif weeks_data["type"] == "interval":
            start = weeks_data["data"]["start"]
            interval = weeks_data["data"]["interval"]
            count = weeks_data["data"]["count"]
            return f"FREQ=WEEKLY;INTERVAL={interval};BYDAY={week_day};COUNT={count}"

        return None
    
    def get_rrule_from_actual_weeks(self, actual_weeks, weekday, start_time):
        """根据实际周次生成 RRULE 规则
        返回 (rrule_string, exdates_list) 元组，如果没有规则则返回 None
        """
        week_day = WEEKDAY_MAP[weekday]
        
        if len(actual_weeks) == 0:
            return None
        
        # 如果周次是连续的，使用 COUNT
        if len(actual_weeks) > 1 and all(actual_weeks[i] == actual_weeks[i-1] + 1 for i in range(1, len(actual_weeks))):
            count = len(actual_weeks)
            return (f"FREQ=WEEKLY;BYDAY={week_day};COUNT={count}", [])
        else:
            # 对于不连续的周次，使用 COUNT 生成从第一个到最后一个周次的所有事件
            # 然后使用 EXDATE 排除不需要的周次
            first_week = actual_weeks[0]
            last_week = actual_weeks[-1]
            count = last_week - first_week + 1
            
            # 计算需要排除的周次
            all_weeks = set(range(first_week, last_week + 1))
            exclude_weeks = sorted(all_weeks - set(actual_weeks))
            
            exdates = []
            for exclude_week in exclude_weeks:
                # 计算排除的日期，使用实际的上课时间
                exclude_date = self.semester_start + timedelta(days=(exclude_week - 1) * 7 + (weekday - 1))
                exclude_dt = datetime.combine(exclude_date, datetime.strptime(start_time, "%H:%M").time())
                exclude_dt = SHANGHAI_TZ.localize(exclude_dt).astimezone(pytz.utc)
                # EXDATE 格式：YYYYMMDDTHHMMSSZ
                exdates.append(exclude_dt.strftime("%Y%m%dT%H%M%SZ"))
            
            rrule = f"FREQ=WEEKLY;BYDAY={week_day};COUNT={count}"
            return (rrule, exdates)

    def write(self, file_path=None):
        """写入 ICS 文件"""
        if not file_path:
            file_path = save_ics_file()
        if not file_path:
            print("用户取消保存，程序退出")
            exit()
        cal = self.generate_ics()
        with open(file_path, "w", encoding='utf-8') as f:
            for line in cal:
                f.write(line.strip() + '\n') if line.strip() else None
                print(f"line: '{line.strip()}'")
        print(f"ICS 文件已保存: {file_path}")
        
def save_ics_file():
    """弹出文件保存窗口，让用户选择保存 ics 路径，并返回文件路径"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)  # 使窗口置顶

    # 适配高 DPI 缩放
    #告诉操作系统使用程序自身的dpi适配
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    #获取屏幕的缩放因子
    ScaleFactor=ctypes.windll.shcore.GetScaleFactorForDevice(0)
    #设置程序缩放
    root.tk.call('tk', 'scaling', ScaleFactor/75)
    
    # 获取当前文件所在目录
    current_dir = os.path.dirname(__file__)

    file_path = filedialog.asksaveasfilename(
        defaultextension=".ics", 
        filetypes=[("iCalendar 文件", "*.ics")],
        initialdir=current_dir
    )
    
    return file_path

if __name__ == "__main__":
    # 示例数据
    data = [
        {
            "course_id": "M402004B",
            "class_id": "03",
            "name": "软件工程",
            "time": {"weekday": 1, "lesson": 1},
            "teacher": "魏名元",
            "location": "逸夫教学楼 YF415",
            "weeks": {"type": "continuous", "data": {"start": 1, "end": 16}},
        },
        {
            "course_id": "C108005B",
            "class_id": "02",
            "name": "概率论与数理统计(B)",
            "time": {"weekday": 3, "lesson": 2},
            "teacher": "刘玉婷",
            "location": "思源楼 SY207",
            "weeks": {"type": "discontinuous", "data": [2, 4, 6, 8, 10, 12, 14, 16]},
        },
        {
            "course_id": "M202006B",
            "class_id": "02",
            "name": "离散数学（A）Ⅱ",
            "time": {"weekday": 5, "lesson": 4},
            "teacher": "王奇志",
            "location": "逸夫教学楼 YF106",
            "weeks": {"type": "interval", "data": {"start": 2, "interval": 2, "count": 8}},
        },
    ]

    # 设定学期起始日期
    semester_start = datetime(2025, 3, 3)

    # 生成并保存 ICS 文件
    writer = Writer(data, semester_start)
    writer.write("timetable.ics")
