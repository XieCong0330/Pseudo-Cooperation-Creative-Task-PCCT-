import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import subprocess
from datetime import datetime
import pandas as pd
import os
def submit():
    name = name_var.get()
    student_id = id_var.get()
    age = age_var.get()
    gender = gender_var.get()

    if not name or not student_id or not age or not gender:
        messagebox.showerror("错误", "请填写所有内容!")
        return

    nowtime = get_current_time_numeric()
    info = [[nowtime],[name, student_id, age, gender]]
    save_to_excel(info,filename=f'{nowtime}.xlsx',sheet_name='subj_info')
    print(f"Name: {name}")
    print(f"Student ID: {student_id}")
    print(f"Age: {age}")
    print(f"Gender: {gender}")

    subprocess.run(["python", "exp_trial.py"])

def get_current_time_numeric():
    now = datetime.now()
    formatted_time = now.strftime("%y%m%d%H")
    return formatted_time

def save_to_excel(data, filename='results.xlsx', sheet_name='Sheet1'):
    """
    将数据保存到 Excel 文件中，每行是一个子列表。

    参数:
    - data: 要保存的数据列表，每个元素是一个子列表，子列表的元素数目可以不相同。
    - filename: 保存的 Excel 文件名，默认为 'results.xlsx'。
    - sheet_name: Excel 文件中的工作表名，默认为 'Sheet1'。
    """
    # 找到最长的子列表的长度
    max_length = max(len(sublist) for sublist in data)

    # 填充每个子列表，使其长度相同
    normalized_data = [sublist + [None] * (max_length - len(sublist)) for sublist in data]

    # 将数据转换为 DataFrame
    df = pd.DataFrame(normalized_data)

    # 检查文件是否存在
    if os.path.exists(filename):
        # 使用 ExcelWriter 进行多工作表写入
        with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    else:
        # 如果文件不存在，创建新文件
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)


# 创建主窗口
root = tk.Tk()
root.title("User Information")
root.geometry("1080x720")

# 使窗口居中
window_width = 1080
window_height = 720

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

position_top = int(screen_height / 2 - window_height / 2)
position_right = int(screen_width / 2 - window_width / 2)

root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

# 创建主框架
frame = ttk.Frame(root)
frame.grid(row=0, column=0, sticky="nsew")

# 设置网格布局权重，使条目在窗口内居中
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)
frame.grid_rowconfigure(0, weight=1)
frame.grid_rowconfigure(1, weight=1)
frame.grid_rowconfigure(2, weight=1)
frame.grid_rowconfigure(3, weight=1)
frame.grid_rowconfigure(4, weight=1)
frame.grid_rowconfigure(5, weight=1)
frame.grid_rowconfigure(6, weight=1)
frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=1)

# 创建变量
name_var = tk.StringVar()
id_var = tk.StringVar()
age_var = tk.StringVar()
gender_var = tk.StringVar()

# 创建并放置指导语
instruction_label = ttk.Label(frame, text="您好，欢迎来到本实验，请填写您的个人信息",font=("Times",30,"bold"))
instruction_label.grid(column=0, row=0, columnspan=2, padx=10, pady=10)

# 创建并放置标签和输入框
name_label = ttk.Label(frame, text="姓名:",font=("Times",30,"bold"))
name_label.grid(column=0, row=1, padx=5, pady=2, sticky="e")
name_entry = ttk.Entry(frame, textvariable=name_var)
name_entry.grid(column=1, row=1, padx=5, pady=2, sticky="w")

id_label = ttk.Label(frame, text="序号:",font=("Times",30,"bold"))
id_label.grid(column=0, row=2, padx=5, pady=2, sticky="e")
id_entry = ttk.Entry(frame, textvariable=id_var)
id_entry.grid(column=1, row=2, padx=5, pady=2, sticky="w")

age_label = ttk.Label(frame, text="年龄:",font=("Times",30,"bold"))
age_label.grid(column=0, row=3, padx=5, pady=2, sticky="e")
age_entry = ttk.Entry(frame, textvariable=age_var)
age_entry.grid(column=1, row=3, padx=5, pady=2, sticky="w")

gender_label = ttk.Label(frame, text="性别:",font=("Times",30,"bold"))
gender_label.grid(column=0, row=4, padx=5, pady=2, sticky="e")
gender_combobox = ttk.Combobox(frame, textvariable=gender_var)
gender_combobox['values'] = ('男', '女')
gender_combobox.grid(column=1, row=4, padx=5, pady=2, sticky="w")

# 创建并放置提交按钮
submit_button = ttk.Button(frame, text="提交", command=submit)
submit_button.grid(column=1, row=5, padx=5, pady=10, sticky="w")

# 运行主循环
root.mainloop()
