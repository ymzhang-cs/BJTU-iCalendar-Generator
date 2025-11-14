from parser import Parser
from ics_writer import Writer
from semester_fetcher import fetch_semester_info

import datetime

if __name__ == "__main__":
    # 尝试自动获取学期信息
    use_auto = input("是否自动从网页获取学期信息？(y/n，默认y): ").strip().lower()
    if not use_auto or use_auto == 'y':
        try:
            semester_start, rest_weeks = fetch_semester_info()
            print(f"\n自动获取成功！")
            print(f"学期开始日期: {semester_start.strftime('%Y-%m-%d')}")
            if rest_weeks:
                rest_info = ", ".join([f"第{week}周后休息{count}周" for week, count in rest_weeks])
                print(f"休息周信息: {rest_info}")
            else:
                print("无休息周")
            
            confirm = input("\n是否使用以上信息？(y/n，默认y): ").strip().lower()
            if confirm and confirm != 'y':
                use_auto = False
        except Exception as e:
            print(f"自动获取失败: {e}")
            print("将使用手动输入模式")
            use_auto = False
    else:
        use_auto = False
    
    # 如果自动获取失败或用户选择手动输入
    if not use_auto:
        semester_start = input("请输入教学周第一周周一的日期（格式：20250224，由于目前不支持节假日推算，若有误差请在此手动调整）：")
        semester_start = datetime.datetime.strptime(semester_start, "%Y%m%d")
        
        # 获取休息周信息
        rest_weeks_input = input("请输入休息周信息（格式：第几周后休息几周，多个用逗号分隔，如：3,1,7,2 表示第3周后休息1周，第7周后休息2周；若无休息周，直接回车）：")
        rest_weeks = []
        if rest_weeks_input.strip():
            parts = rest_weeks_input.split(",")
            if len(parts) % 2 != 0:
                print("错误：休息周信息格式不正确，应为：第几周后休息几周，多个用逗号分隔")
                exit()
            for i in range(0, len(parts), 2):
                try:
                    after_week = int(parts[i].strip())
                    rest_count = int(parts[i + 1].strip())
                    rest_weeks.append((after_week, rest_count))
                except ValueError:
                    print(f"错误：无法解析休息周信息：{parts[i]}, {parts[i + 1]}")
                    exit()
            # 按 after_week 排序
            rest_weeks.sort(key=lambda x: x[0])
    
    print("\n正在生成课表...")

    parser = Parser()
    data = parser.parse()
    
    print("课表解析成功！")
    
    writer = Writer(data, semester_start, rest_weeks)
    writer.write()
    print("课表生成成功！")
    
    
    