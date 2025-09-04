import tkinter as tk
from tkinter import filedialog
import ctypes
from bs4 import BeautifulSoup
import re
import os

class Parser:
    def __init__(self, file_path=None):
        print("file_path:", file_path if file_path else "None")
        self.file_path = file_path if file_path else select_html_file()
        if not self.file_path:
            print("未选择文件，程序退出")
            exit()
        print("self.file_path:", self.file_path)

    def parse(self):
        """
        解析课表 HTML 文件，返回解析后的数据
        数据格式：
        [
            {
                "course_id": "M402004B",
                "class_id": "03",
                "name": "软件工程",
                "time": {
                    "weekday": 1,
                    "lesson": 1,
                },
                "teacher": "魏名元",
                "location": "逸夫教学楼 YF415",
                "weeks": {
                    "type": "continuous",
                    "data": {
                        "start": 1,
                        "end": 16,
                    },
                }
                // OR
                "weeks": {
                    "type": "discontinuous",
                    "data": [2, 4, 6, 8, 10, 12, 14, 16],
                }
                // OR
                "weeks": {
                    "type": "interval",
                    "data": {
                        "start": 1,
                        "interval": 2,
                        "count": 8,
                    },
                }
            }
        ]
        """
        
        # 解析后的数据
        parsed_data = []
        
        with open(self.file_path, "r", encoding="utf-8") as f:
            html = f.read()
        
        # 使用 BeautifulSoup 解析 HTML
        soup = BeautifulSoup(html, "html.parser")
        
        # 选择表格：<table class="table table-bordered">
        table = soup.find("table", class_="table table-bordered")
        
        # 选择表格中的所有行：<tr>
        rows = table.find_all("tr")
        
        # 遍历每一行 tr
        # 第一行是表头，不需要解析
        for lesson_idx, row in enumerate(rows[1:], start=1):
            # 选择行中的所有单元格：<td>
            cells = row.find_all("td", recursive=False)
            
            # 遍历每一个单元格 td
            # 第一个单元格是时间，不需要解析
            for weekday_idx, cell in enumerate(cells[1:], start=1):
                
                # 遍历每一个div
                for div in cell.find_all("div", recursive=False):
                    """
                    示例数据：
                    <div >
                        <span >
                            M402004B [03] <br />
                            软件工程<br />
                        </span>

                        <div style="max-width:120px;">
                            第01-16周
                            <i>魏名元</i>
                        </div>
                        <span class="text-muted">海淀西校区, 逸夫教学楼, YF东706</span>
                        <span class="green" style="display: inline-block;">[ 选中 ]</span>
                    </div>
                    """
                    print(div)
                    
                    # 解析课程信息
                    course_info = div.find("span").get_text().strip()
                    course_id, class_id, name = re.match(r"(\w+) \[(\w+)\]\s+(.+)", course_info).groups()
                    
                    # 解析上课周数和老师
                    week_teacher_info = div.find_all("div")[0]
                    weeks_str, teacher_str = week_teacher_info.get_text().strip().split("\n")
                    # 去除老师前面的空格
                    teacher = teacher_str.strip()
                    # 识别周数格式
                    time_type, time_data = week_type_detect(weeks_str)
                    
                    # 解析上课地点
                    # 20250905 修改：发现学校教务系统添加了校区信息，并且分隔使用", " 这里暂时只取第2个和第3个
                    location = div.find_all("span")[1].get_text().strip().split(", ")
                    location = location[1] + " " + location[2]
                    
                    # 保存解析后的数据
                    parsed_data.append({
                        "course_id": course_id,
                        "class_id": class_id,
                        "name": name,
                        "time": {"weekday": weekday_idx, "lesson": lesson_idx},
                        "teacher": teacher,
                        "location": location,
                        "weeks": {"type": time_type, "data": time_data}
                    })

        return parsed_data
    
def week_type_detect(weeks_str):
    """
    判断周数格式，并返回 time_type 和 time_data
    time_type: "continuous", "discontinuous", "interval"
    time_data:
        - continuous: {"start": int, "end": int}
        - discontinuous: [int, int, ...]
        - interval: {"start": int, "interval": int, "count": int}
    """
    
    if "-" in weeks_str:
        start_week, end_week = re.match(r"第(\d+)-(\d+)周", weeks_str).groups()
        time_type = "continuous"
        time_data = {"start": int(start_week), "end": int(end_week)}
    elif ", " in weeks_str:
        weeks = re.match(r"第(.+)周", weeks_str).groups()[0].split(", ")
        time_type = "discontinuous"
        time_data = [int(week) for week in weeks]
        # 进一步判断是否为间隔周数
        if len(time_data) > 2:
            interval = time_data[1] - time_data[0]
            if all(time_data[i] - time_data[i-1] == interval for i in range(1, len(time_data))):
                time_type = "interval"
                time_data = {"start": time_data[0], "interval": interval, "count": len(time_data)}
    else:
        raise ValueError(f"Unknown week format: {weeks_str}")
    
    return time_type, time_data

def select_html_file():
    """弹出文件选择窗口，让用户选择课表 HTML 文件，并返回文件路径"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
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
    # + "pages"
    pages_dir = os.path.join(current_dir, "pages")

    # 弹出文件选择窗口
    file_path = filedialog.askopenfilename(
        title="选择课表 HTML 文件",
        filetypes=[("HTML 文件", "*.html"), ("所有文件", "*.*")],
        initialdir=pages_dir
    )

    return file_path  # 返回选择的文件路径

if __name__ == "__main__":
    parser = Parser()
    data = parser.parse()
    for lesson in data:
        print(lesson)
