import pandas as pd

# 读取模块 - 标准化后的选课时段数据
class readData:
    def __init__(self):
        # 标准化后的选课数据目录
        self.data_dir = "./data.excel"
        # 表1为Slot数据
        # 表2为Counseling对应的时间
        # 表3为课程名称标准 对应教师 对应老师
        self.bind = pd.DataFrame()
        self.math = pd.DataFrame()
        self.math_bind = pd.DataFrame()
        self.classes_info = pd.DataFrame()
        self.student_grade = pd.DataFrame()
        self.slots = ["Slot 1", "Slot 2", "Slot 3", "Slot 4", "Slot 5", "Slot 6", "Slot 7"]

    def readData(self, directory):
        try:
            self.bind = pd.read_excel(directory, sheet_name="选课绑定")
            self.math = pd.read_excel(directory, sheet_name="数学课")
            self.math_bind = pd.read_excel(directory, sheet_name="数学绑定")
            self.classes_info = pd.read_excel(directory, sheet_name="课程信息")
            self.student_idname = pd.read_excel(directory, sheet_name="学号姓名")
            # 预处理数据
            self.doData()
            return True
        except:
            return False

    def doData(self):
        self.bind = self.bind.iloc[:5]
        self.math = self.math["Math"].tolist()
        '''
        # 获取学号对应的班级
        for index_student_class in range(len(self.student_grade["班级"])):
            self.student_grade.loc[index_student_class, "班级"] = self.student_grade["班级"][index_student_class][0:3]
        '''

