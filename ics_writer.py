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
    def __init__(self, data, semester_start):
        """
        :param data: 课程数据列表
        :param semester_start: 学期开始日期 (datetime 类型)
        """
        self.data = data
        self.semester_start = semester_start  # 例如 datetime(2025, 3, 3)

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

            # 计算课程首次上课日期
            first_week = self.get_first_week(weeks_data)
            first_class_date = self.semester_start + timedelta(days=(first_week - 1) * 7 + (weekday - 1))
            start_dt = datetime.combine(first_class_date, datetime.strptime(start_time, "%H:%M").time())
            
            # 转换为 Asia/Shanghai 时区
            start_dt = SHANGHAI_TZ.localize(start_dt).astimezone(pytz.utc)
            end_dt = start_dt + timedelta(minutes=110 if lesson != 7 else 50)


            event = Event()
            event.name = f"{course_name} - {teacher}"
            event.begin = start_dt
            event.end = end_dt
            event.location = location

            # 生成 RRULE
            rrule = self.get_rrule(weeks_data, weekday)
            if rrule:
                event.extra.append(ContentLine(name="RRULE", value=rrule))  # ✅ 这里使用 ContentLine

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
        """生成 RRULE 规则"""
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

    def write(self, file_path=None):
        """写入 ICS 文件"""
        if not file_path:
            file_path = save_ics_file()
        if not file_path:
            print("用户取消保存，程序退出")
            exit()
        cal = self.generate_ics()
        with open(file_path, "w", encoding='utf-8') as f:
            f.writelines(cal)
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
