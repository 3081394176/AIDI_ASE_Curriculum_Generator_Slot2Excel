为爱迪学校美国高中部设计的根据学生选课时段生成最终课表项目

需要的库
pip3 install pandas openpyxl matplotlib customtkinter
python版本 3.11

Windows编译
pyinstaller --onefile --windowed --icon=icon.ico main.py

v1.0
完成了根据Slot信息生成excel版本课表

v1.1
适配了新版课表模版 (包含new的Excel名字统一视为新版课表)

v1.2
增加了体育课数学检查，在data中的数学绑定表格增加了强制的体育课时段对应数据