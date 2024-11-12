import os
import shutil

# 文件夹路径，存放xlsx文件
folder_path = "/Users/a3081394176/Desktop/schedules_finished_in_2024_9_2_18_23_18"

# 班级列表
class_list = ["Y10-1", "Y10-2", "Y10-3"]  # 这里替换为你的班级列表

# 创建班级文件夹并移动文件
for class_name in class_list:
    # 创建班级文件夹路径
    class_folder = os.path.join(folder_path, class_name)

    # 如果班级文件夹不存在，则创建
    if not os.path.exists(class_folder):
        os.makedirs(class_folder)

    # 遍历文件夹中的所有文件
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx") and class_name in filename:
            # 构建源文件路径和目标文件路径
            source_file = os.path.join(folder_path, filename)
            destination_file = os.path.join(class_folder, filename)

            # 移动文件
            shutil.move(source_file, destination_file)
            print(f"Moved {filename} to {class_folder}")
