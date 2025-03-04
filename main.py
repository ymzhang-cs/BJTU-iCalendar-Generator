from parser import Parser
from ics_writer import Writer

import datetime

if __name__ == "__main__":
    semester_start = input("请输入教学周第一周周一的日期（格式：20250224）：")
    semester_start = datetime.datetime.strptime(semester_start, "%Y%m%d")
    
    print("正在生成课表...")

    parser = Parser()
    data = parser.parse()
    
    print("课表解析成功！")
    
    writer = Writer(data, semester_start)
    writer.write()
    print("课表生成成功！")
    
    
    