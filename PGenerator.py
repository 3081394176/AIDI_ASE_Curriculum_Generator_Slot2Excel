import pandas as pd
from Drawer import Drawer
import os
from datetime import datetime

class PGenerator:

    def __init__(self):
        return

    def __init__(self, read, xuanke_excel_directory, schedule_format_directory, output_directory, format, platform):
        self.current_student = -1
        self.students_num = -1
        self.student_current_index = -1
        self.read = read
        self.xuanke_excel_directory = xuanke_excel_directory
        self.schedule_format_directory = schedule_format_directory
        self.xuanke_excel = pd.DataFrame()
        self.output_directory = output_directory
        self.format = format
        self.platform = platform
        self.slots = ["Slot 1", "Slot 2", "Slot 3", "Slot 4", "Slot 5", "Slot 6", "Slot 7"]
        self.kebiao = {
            'MON': [None] * 7,
            'TUE': [None] * 7,
            'WED': [None] * 7,
            'THU': [None] * 7,
            'FRI': [None] * 7
        }
        # 设置课表
        self.kebiao_df = pd.DataFrame(self.kebiao)
        self.room_df = pd.DataFrame()

        self.student_id = ""
        self.student_name = ""
        self.student_class = ""

        # 初始化输出目录
        self.createFolfer()

    def kebiaoReset(self):
        self.kebiao_df = pd.DataFrame(self.kebiao)

    def readStudentCourseExcel(self):
        # 读取学生选课数据
        try:
            self.xuanke_excel = pd.read_excel(self.xuanke_excel_directory)
            self.students_num = len(self.xuanke_excel)
            return True
        except:
            return False

    # 创建用于存储导出数据的文件夹
    def createFolfer(self):
        try:
            # 创建单个文件夹
            folder_path = f"{self.output_directory}/schedules_finished_in_{datetime.now().year}_{datetime.now().month}_{datetime.now().day}_{datetime.now().hour}_{datetime.now().minute}_{datetime.now().second}/"
            os.makedirs(folder_path, exist_ok=True)
            self.output_directory = folder_path
        except:
            raise Exception("导出目录文件夹生成失败")

    def doAllClasses(self):
        for student_index in range(len(self.xuanke_excel)):
            # 处理非体育时段的课程并且修改kebiao_df
            self.student_id = self.xuanke_excel.iloc[student_index]["学号"]
            self.student_current_index = student_index
            # 正在处理的学生
            self.current_student = str(student_index) + ' - ' + str(self.student_id)
            #self.student_name = self.xuanke_excel.iloc[student_index]["姓名"]

            try:
                self.student_name = str(
                    self.read.student_idname[self.read.student_idname['学号'] == self.student_id]['姓名'].tolist()[0])
            except:
                raise Exception(f'未找到学号对应的姓名: {self.student_id}')

            try:
                self.student_class = str(
                    self.read.student_idname[self.read.student_idname['学号'] == self.student_id]['班级'].tolist()[0])
            except:
                raise Exception(f'未找到学号对应的班级: {self.student_id}')

            slot_classes = self.xuanke_excel.iloc[student_index][["sot1", "sot2", "sot3", "sot4", "sot5", "sot6", "sot7"]].tolist()

            # 生成正常课程循环
            for slot in self.slots:
                for slot_index in range(5):
                    if slot_classes[int(slot[5:6]) - 1] != "体育" and slot_classes[int(slot[5:6]) - 1] != "PE":
                        # self.kebiao_df[self.read.sheet1[slot][slot_index][0:3]][int(self.read.sheet1[slot][slot_index][3:4]) - 1] = slot
                        self.kebiao_df.loc[int(self.read.bind[slot][slot_index][3:4]) - 1, self.read.bind[slot][slot_index][0:3]] = slot

            # 处理体育时段: Counseling + 体育 + 数学

            # 获取数学时段
            math_slot = -1

            for course_index in range(len(slot_classes)):
                if (slot_classes[course_index] in self.read.math) & (math_slot == -1):
                    math_slot = course_index + 1
                    break

            # print("math_slot" + str(math_slot))

            #获取年级
            #grade = self.read.student_grade["班级"][self.read.student_grade.index[self.read.student_grade['学号'] == self.student_id].tolist()].tolist()[0]
            grade = self.student_class[0:3]

            # math_slot数学时段
            counseling_cate = ''

            if (math_slot == 1):
                counseling_cate = 'A'
            elif (math_slot == 2):
                counseling_cate = 'B'
            elif (math_slot == 3):
                counseling_cate = 'C'
            elif (math_slot == 4):
                counseling_cate = 'D'

            # 根据绑定修改课表信息
            for math_bind_index in range(len(self.read.math_bind)):
                if (int(self.read.math_bind["Slot"][math_bind_index][5:6]) == math_slot):

                    self.kebiao_df.loc[
                        int(self.read.math_bind["Mathbind1"][math_bind_index][3:4]) - 1,
                        self.read.math_bind["Mathbind1"][math_bind_index][0:3]] = self.read.math_bind["Slot"][math_bind_index]

                    self.kebiao_df.loc[
                        int(self.read.math_bind["Mathbind2"][math_bind_index][3:4]) - 1,
                        self.read.math_bind["Mathbind2"][math_bind_index][0:3]] = self.read.math_bind["Slot"][
                        math_bind_index]

                    self.kebiao_df.loc[
                        int(self.read.math_bind["Counseling"][math_bind_index][3:4]) - 1,
                        self.read.math_bind["Counseling"][math_bind_index][0:3]] = grade + " Counseling " + counseling_cate

                    break

            # 填补None为体育
            self.kebiao_df = self.kebiao_df.fillna("PE")

            # 替换时段为课程名称
            for day in ["MON", "TUE", "WED", "THU", "FRI"]:
                for kebiao_index in range(7):
                    # print(self.kebiao_df.loc[kebiao_index, day])
                    if not (self.kebiao_df.loc[kebiao_index, day] == "PE" or self.kebiao_df.loc[kebiao_index, day] == "体育" or self.kebiao_df.loc[kebiao_index, day] == grade + " Counseling " + counseling_cate):
                        self.kebiao_df.loc[kebiao_index, day] = slot_classes[int(self.kebiao_df.loc[kebiao_index, day][5:6]) - 1]

            # 获取替换dict

            class_dict = dict(zip(self.read.classes_info['系统课程名称'].tolist(), self.read.classes_info['课表课程名称'].tolist()))

            print(class_dict)

            # 替换课程名称

            self.kebiao_df.replace(class_dict, inplace=True)

            print(self.kebiao_df)

            # 获取课程教室 然后生成同课表结构但是内容为教室的课表
            slot_classes = [class_dict.get(item, item) for item in self.xuanke_excel.iloc[student_index][
                ["sot1", "sot2", "sot3", "sot4", "sot5", "sot6", "sot7"]].tolist()]

            classes_place = dict(zip(slot_classes, [
                self.read.classes_info[self.read.classes_info['课表课程名称'] == x].reset_index()[
                    f'时段{slot_classes.index(x) + 1}'].iloc[0]
                if not self.read.classes_info[self.read.classes_info['课表课程名称'] == x].empty
                else None  # 或者设置其他默认值
                for x in slot_classes
            ]))

            print(classes_place)

            self.room_df = self.kebiao_df.copy()
            self.room_df.replace(classes_place, inplace=True)
            self.room_df.replace(grade + ' Counseling ' + counseling_cate, self.read.classes_info[self.read.classes_info['课表课程名称'] == grade + " Counseling " + counseling_cate].reset_index()[
                    f'时段{1}'].iloc[0] if not self.read.classes_info[self.read.classes_info['课表课程名称'] == (grade + " Counseling " + counseling_cate)].empty else 'Counseling', inplace=True)

            print(self.room_df)

            # 生成当前学生课表
            # 初始化绘制class
            drawer = Drawer(self.output_directory, self.student_name, self.student_id, self.student_class, self.kebiao_df, self.room_df, self.schedule_format_directory, self.platform)
            # 调用绘制函数绘制当前课表
            drawer.Draw(self.format)

            # 重置课表
            print(slot_classes)
            self.kebiaoReset()






