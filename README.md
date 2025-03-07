# BJTU-iCalendar-Generator

![Visitors](https://api.visitorbadge.io/api/visitors?path=https%3A%2F%2Fgithub.com%2Fymzhang-cs%2FBJTU-iCalendar-Generator&countColor=%23263759)

将教务系统的课程表导出为 iCalendar 文件，方便导入到日历软件中。

支持以下功能：

- [x] 课程名称、节次、地点、教师、周数等信息的提取
- [x] 课程循环周期识别，设定重复日程
- [x] 错峰上课时间识别
- [x] 友好的文件选择界面

未来可能支持：

- [ ] 日程名称显示格式的自定义
- [ ] 其他日历软件操作指引
- [ ] 邮件发送功能或在线服务

## 使用方法

1. 安装所需模块

```bash
pip install -r requirements.txt
```

2. 登录教务系统，进入课表页面。

3. `Ctrl + S` 或 右键 -> 另存为，设置保存格式为 `网页，仅 HTML (*.html;*.htm)`，保存到该项目的 `pages` 文件夹下。

4. 运行 `main.py`，按照操作指引选择 html 文件与导出的 iCalendar 文件的路径。

5. 导入 iCalendar 文件到日历软件中。

## ICS 文件导入 iOS 日历

1. 使用邮箱发送 iCalendar 文件到自己的邮箱（需绑定到原生的邮件 App）

2. 在邮件 App 中打开 iCalendar 文件，点击“添加全部”，选择要保存到的日历（推荐新建一个），即可完成导入。

## FAQ

### 课程名称/时间/地点/周期识别不正确

请检查教务系统导出的课表是否正确，或者提交 issue。
