import tkinter as tk
from tkinter import filedialog, messagebox, Tk
import customtkinter

from readData import readData
from PGenerator import PGenerator

import platform

version = '1.0'
software_name = 'Cactus'

class ASEScheduleGeneratorApp:
    def __init__(self):
        # 初始化隐藏的tk
        # self.root = Tk()
        # self.root.withdraw()

        # 获取系统信息
        self.platform = self.get_os_type()

        # 初始化 CustomTkinter 设置
        customtkinter.set_appearance_mode("system")  # Modes: system (default), light, dark
        self.app = customtkinter.CTk()
        self.app.geometry("400x280")
        self.app.resizable(False, False)
        self.app.title(f'ASE - {software_name} v{version} - ' + self.platform)

        # 初始化变量
        self.initialization_status = False
        self.read = readData()
        self.format = 'Excel'
        self.supported_format = ['Excel', '图片', 'PDF']

        # 初始化界面
        self.create_widgets()



    def get_os_type(self):
        system = platform.system()
        if system == "Windows":
            return "Windows"
        elif system == "Darwin":
            return "macOS"
        elif system == "Linux":
            return "Linux"
        else:
            return "Unknown"

    def create_widgets(self):
        # 配置行列权重，使所有组件居中对齐
        self.app.grid_columnconfigure(0, weight=1)

        # 创建标题
        title = customtkinter.CTkLabel(self.app, text=f'ASE - {software_name} v{version}', font=customtkinter.CTkFont(size=20, weight="bold"))
        title.grid(row=0, column=0, padx=20, pady=(20, 10), sticky='nsew')

        # 创建初始化按钮
        self.button1 = customtkinter.CTkButton(master=self.app, text="初始化", command=self.on_button1_click)
        self.button1.grid(row=1, column=0, padx=20, pady=(10, 10), sticky='nsew')

        # 创建生成课表按钮
        self.button2 = customtkinter.CTkButton(master=self.app, text="生成课表", command=self.on_button2_click)
        self.button2.grid(row=4, column=0, padx=20, pady=(10, 10), sticky='nsew')

        # 创建导出格式标签
        supported_format_label = customtkinter.CTkLabel(self.app, text="导出格式:", anchor="w")
        supported_format_label.grid(row=2, column=0, padx=20, pady=(10, 0), sticky='nsew')

        # 创建格式选择菜单
        supported_format_optionmenu = customtkinter.CTkOptionMenu(self.app, values=self.supported_format, command=self.set_format)
        supported_format_optionmenu.grid(row=3, column=0, padx=20, pady=(10, 10), sticky='nsew')

    def set_format(self, selected_format: str):
        self.format = selected_format
        print(self.format)

    def on_button1_click(self):
        print('按钮1')
        directory = filedialog.askopenfilename()
        if self.read.readData(directory):
            messagebox.showinfo("提示", "数据加载成功")
            self.initialization_status = True
            self.button1.configure(state=tk.DISABLED)  # 禁用按钮
        else:
            messagebox.showinfo("提示", "数据加载失败，请选择数据excel")

    def on_button2_click(self):
        print('按钮2')
        if not self.initialization_status:
            messagebox.showinfo("提示", "请先进行初始化")
            return

        messagebox.showinfo("提示", "请选择学生选课Excel")
        xuanke_excel_directory = filedialog.askopenfilename()

        messagebox.showinfo("提示", "请选择课表模版Excel")
        schedule_format_directory = filedialog.askopenfilename()

        messagebox.showinfo("提示", "请选择输出目录")
        output_directory = filedialog.askdirectory()

        if not ("xlsx" in xuanke_excel_directory or "xls" in xuanke_excel_directory):
            messagebox.showinfo("提示", "错误目录，请重新选择")
        else:
            try:
                generator = PGenerator(self.read, xuanke_excel_directory, schedule_format_directory, output_directory, self.format, self.platform)
                generator.readStudentCourseExcel()
                generator.doAllClasses()
                messagebox.showinfo("提示", "课表生成完成！")
            except Exception as e:
                try:
                    current = str(generator.current_student)
                except Exception as ee:
                    messagebox.showerror("错误", f"课表生成过程中发生错误: {ee}\n")
                messagebox.showerror("错误", f"课表生成过程中发生错误: {e}\n" + str(generator.current_student))

    def run(self):
        self.app.mainloop()

# 主程序入口
if __name__ == "__main__":
    try:
        app = ASEScheduleGeneratorApp()
        app.run()
    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {e}")
