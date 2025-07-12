import tkinter as tk
from tkinter import messagebox
import random
import pandas as pd
import networkx as nx
import requests
import time
from datetime import datetime
import os
# 准备词汇列表，每个词只列一次

base_words = ["砖头","便利贴","筷子","钥匙","报纸"]
random.shuffle(base_words)
# 先生成所有词的联想任务
# Generating tasks as dictionaries
association_tasks = [[word, "联想", 5000] for word in base_words for _ in range(10)]
pattern = [(1, 1), (1, 1), (1, -1), (1, -1), (-1, 1), (-1, 1), (-1, -1), (-1, -1)]

# Generating creativity tasks with varying last two elements based on the specified pattern
creativity_tasks = []
for word in base_words:
    random.shuffle(pattern)
    for i in range(8):  # Create 8 entries per word
        values = pattern[i]  # Select values from the pattern based on index
        creativity_tasks.append([word, "创造性", 12000] + list(values))

org_list = []
timers = {}
# Randomize the tasks
random.shuffle(association_tasks)
random.shuffle(creativity_tasks)


tasks = association_tasks+creativity_tasks
print(tasks)
# 读取示例


# 创建主窗口
root = tk.Tk()
root.title("心理学实验")

# 设置窗口大小
root.geometry("1080x720")
window_width = 1080
window_height = 720

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

position_top = int(screen_height / 2 - window_height / 2)
position_right = int(screen_width / 2 - window_width / 2)

root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
# 创建用于显示任务的标签
display_label = tk.Label(root, text="你好", font=("Arial", 24),wraplength=800)
display_label.pack(expand=True)

# 创建输入框供被试输入
entry = tk.Entry(root, font=("Times", 30))
entry.pack(pady=20)


# 创建提交按钮
submit_button = tk.Button(root, text="提交", command=lambda: submit_association())
submit_button.pack(pady=10)
continue_button = tk.Button(root, text="继续", command=lambda: prepare_trial())
continue_button.pack(pady=20)
rate_button = tk.Button(root, text="评价观点结束", command=lambda: show_examp_rate())
rate_button.pack(pady=20)
sele_button = tk.Button(root, text="评价观点结束", command=lambda: show_examp_sele())
sele_button.pack(pady=20)
submit_sele_button = tk.Button(root, text="关联评价结束", command=lambda: submit_sele())
submit_sele_button.pack(pady=20)

continueAUT_button = tk.Button(root, text="继续", command=lambda: startAUT_exp())
continueAUT_button.pack(pady=20)

task_index = 0
association_list = []
AUT_list = []
example_list =[]
accpt_idea =[]
# 提交联想或创造性用途
rating_var = tk.IntVar()  # This variable will store the user's rating
rating_frame = tk.Frame(root)
tk.Label(rating_frame, text="1为非常低，5为非常高:", font=("Times",30)).pack()
result_list = []
# Create and pack radio buttons within the frame
for i in range(1, 6):
    tk.Radiobutton(rating_frame, text=str(i), variable=rating_var, value=i, font=("Times", 30)).pack(side=tk.LEFT)

def start_exp():
    entry.pack_forget()
    submit_button.pack_forget()
    continueAUT_button.pack_forget()
    continue_button.pack()
    rate_button.pack_forget()
    sele_button.pack_forget()
    submit_sele_button.pack_forget()
    display_label.config(text=f"您好，欢迎来到本实验，本实验分为两个部分，第一部分为自由联想任务，"
                              f"需要您看见提示词后，以最快的速度联想该提示词的一个相关的词语，如：夏天联想到西瓜，"
                              f"联想结束后请按下空格键，并输入您的联想内容。注意，请不要联想特定的名字(如：马克思)"
                              f"准备好后请点击‘继续’",font=("Times",30))  # 显示当前任务

def startAUT_exp():
    entry.pack_forget()
    submit_button.pack_forget()
    continueAUT_button.pack_forget()
    continue_button.pack()
    rate_button.pack_forget()
    sele_button.pack_forget()
    submit_sele_button.pack_forget()
    display_label.config(text=f"您好，接下来为本实验的第二部分，第二部分为多用途任务，"
                              f"在该部分，您需要与合作者一起思考某个物品的创造性用途，如：小刀的用途是开锁，"
                              f"您会首先得到合作者的答案，在合作者答案的基础上进行思考答案，思考结束后请按下空格键，并输入您的答案。"
                              f"准备好后请点击‘继续’。",font=("Times",30))  # 显示当前任务
# 准备每个试次
def prepare_trial():
    global result
    entry.pack_forget()
    submit_button.pack_forget()
    continue_button.pack_forget()
    rate_button.pack_forget()
    sele_button.pack_forget()
    submit_sele_button.pack_forget()
    result = []
    if task_index < len(tasks):
       display_label.config(text="+", fg="black",font=("Times",40,"bold"))  # 显示注视点
    # 设置延迟后显示任务
       delay = random.randint(1000, 3000)
       if tasks[task_index][1] == '创造性':
           root.after(delay, lambda: show_example(tasks[task_index]))
       else:
           root.after(delay, lambda: show_task(tasks[task_index]))
    else:
        finish_experiment()

def show_example(task):
    entry.pack_forget()
    submit_button.pack_forget()
    continue_button.pack_forget()
    rating_frame.pack_forget()
    judge_len = [item for item in association_list if item[0].strip() == tasks[task_index][0]]
    ans_len = [item for item in AUT_list if item[0].strip() == tasks[task_index][0]]
    exam_len = [item for item in accpt_idea if item[0].strip() == tasks[task_index][0]]



    judge_parts = [sublist[1] for sublist in judge_len if len(sublist) > 1]
    ans_parts = [sublist[1] for sublist in ans_len if len(sublist) > 1]
    exam_parts = [sublist[1] for sublist in exam_len if len(sublist) > 1]

    # 合并这些部分到一个新列表中
    combined_list = [tasks[task_index][0]] + judge_parts + ans_parts + exam_parts
    print(combined_list)

    # 检查合并后的列表是否为空
    if len(combined_list) > 1:
        # combined_list 包含任务名称和至少一个有效元素
        final_list = combined_list
    else:
        final_list = None  # 或者提供一个默认的列表或错误处理
    path_org, cluster_org = analyze_modified_network(final_list)
    print(path_org, cluster_org)
    data = pd.read_excel('new_AUT3.xls', sheet_name=tasks[task_index][0])

    # 将DataFrame转换为列表
    # 如果你只想转换某个特定列，可以使用 data['column_name'].tolist()
    data = data.iloc[:, 0]
    data_list = data.values.tolist()
    random.shuffle(data_list)
    result_exp = find_matching_result(data_list, final_list, tasks, task_index, path_org, cluster_org)
    result.append(result_exp[1])
    result.append(result_exp[2])
    result.append(result_exp[3])
    example_list.append(result_exp)
    example = result_exp[1]
    print(example)
    display_label.config(text=f"{task[0]}的用途是： {example} ", fg="black",font=("Times",30))  # 显示当前任务
    root.after(5000, lambda: show_task(tasks[task_index]))
# 显示当前任务
def show_task(task):
    global start_time
    continue_button.pack_forget()
    display_label.config(text=f" {task[1]} '{task[0]}'，准备好后按空格", fg="black",font=("Times",30))  # 显示当前任务
    root.bind('<space>', lambda event: space_pressed(task))
    # 设置一个计时器，并保存ID以便可以取消
    timer_id = root.after(task[2], lambda: space_pressed(task, auto_triggered=True))
    # 将timer_id存储在task元组中，以便后续可以取消
    task.append(timer_id)
    result.append(task[1])
    result.append(task[0])
    start_time = time.time()
    print(task)



def space_pressed(task, auto_triggered=False):
    global start_time
    global end_time
    end_time = time.time()
    continue_button.pack_forget()
    entry.delete(0, tk.END)  # 清空输入框
    # 检查是否自动触发，如果是由于计时结束自动触发，则不取消计时器（因为已经触发）
    if not auto_triggered and len(task) > 3:
        root.after_cancel(task[-1])  # 取消之前设置的计时器
        task.pop()  # 移除存储的timer_id
    elapsed_time = end_time - start_time
    result.append(elapsed_time)
    print(elapsed_time)
    root.unbind('<space>')  # 取消绑定空格键事件，避免重复触发
    display_label.config(text="请输入您的答案:",font=("Times",30))
    entry.pack(pady=20)
    submit_button.pack(pady=10)
    entry.focus_set()  # 自动聚焦到输入框

def submit_association():
    global task_index
    submit_time = time.time()
    elapsed_time = submit_time - end_time
    result.append(elapsed_time)

    print(elapsed_time)
    user_input = entry.get()  # 获取输入
    result.append(user_input)
    print(f"用户任务：{tasks[task_index][1]} - '{tasks[task_index][0]}' - 联想到：{user_input}")
    sim = calculate_similarity(tasks[task_index][0], user_input)
    if sim == None:
        messagebox.showinfo("无法理解", "无法理解该答案，请重新作答。")  # 显示提示信息
        entry.pack_forget()
        submit_button.pack_forget()
        show_task(tasks[task_index])
        return
    if tasks[task_index][1] == '联想':

        judge_len = [item for item in association_list if item[0].strip() == tasks[task_index][0]]
        words = [item[1] for item in judge_len]
        if user_input in words:
            messagebox.showinfo("重复答案", "该答案已重复，请重新作答。")  # 显示提示信息
            entry.pack_forget()
            submit_button.pack_forget()
            show_task(tasks[task_index])
            return
        else:
            association_list.append([tasks[task_index][0], user_input, sim])
            judge_len = [item for item in association_list if item[0].strip() == tasks[task_index][0]]
            words = [item[1] for item in judge_len]
        print(words)

        if len(words) == 10:

            words.insert(0, tasks[task_index][0])
            path_len, clustering_coefficient = analyze_modified_network(words)
            org = [tasks[task_index][0], path_len, clustering_coefficient]
            org_list.append(org)
            result.append(path_len)
            result.append(clustering_coefficient)
            result_list.append(result)

            print(org_list)
        else:
            result_list.append(result)

    else:
        AUT_list.append([tasks[task_index][0], user_input, sim])


    task_index += 1
    if task_index == len(association_tasks):
        show_rest_screen()
    elif task_index > len(association_tasks):
        show_rating_screen()
    elif task_index >= len(tasks):
        finish_experiment()
    else:
        prepare_trial()





# 结束实验
def show_rating_screen():
    entry.pack_forget()
    submit_button.pack_forget()
    rating_var.set(0)

    display_label.config(text=f"您的答案是：{entry.get()}，请问您认为该答案的创造性水平为何", fg="black",font=("Times",30))
    rating_frame.pack(pady=20)  # Show the rating frame for creativity level

    rate_button.pack(pady=20)

def show_examp_rate():
    entry.pack_forget()
    submit_button.pack_forget()
    rate_button.pack_forget()
    self_rate = rating_var.get()
    result.append(self_rate)
    rating_var.set(0)
    example = example_list[task_index-len(association_tasks)-1][1]

    display_label.config(text=f"他人的答案是：{example}，请问您认为该答案的创造性水平为何", fg="black",font=("Times",30))

    sele_button.pack(pady=20)

def show_examp_sele():
    entry.pack_forget()
    submit_button.pack_forget()
    sele_button.pack_forget()
    self_rate = rating_var.get()
    result.append(self_rate)
    rating_var.set(0)


    display_label.config(text=f"请问您认为该答案和您的答案的关系的密切程度是", fg="black", font=("Times", 30))

    submit_sele_button.pack(pady=20)

def submit_sele():
    rating_frame.pack_forget()
    self_rate = rating_var.get()
    result.append(self_rate)
    example = example_list[task_index - len(association_tasks) - 1][1]
    accpt_idea.append(example)
    entry.delete(0, tk.END)
    prepare_trial()
def show_rest_screen():
    entry.pack_forget()
    submit_button.pack_forget()
    continueAUT_button.pack()
    display_label.config(text="休息一下, 按 '继续' 结束休息", fg="black",font=("Times",30,"bold"))

def finish_experiment():
    display_label.config(text="实验结束，请联系主试", fg="black",font=("Times",30,"bold"))
    entry.pack_forget()
    submit_button.pack_forget()
    continue_button.pack_forget()

    save_to_excel(result_list, filename=f'{numeric_time}.xlsx',sheet_name='data')

# 计算语义距离与语义网络
def get_text_vector(text, server_url='http://127.0.0.1:5000/vectorize_text'):
    """
    获取文本的词向量表示，通过调用 Flask 服务
    """
    response = requests.post(server_url, json={'text': text})
    if response.status_code == 200:
        data = response.json()
        return data.get('vector')
    else:
        return None  # 或者处理错误

def calculate_similarity(text1, text2, server_url='http://127.0.0.1:5000/similarity'):
    """
    通过 Flask 服务计算两个文本之间的语义距离
    """
    vec1 = get_text_vector(text1)
    vec2 = get_text_vector(text2)
    if vec1 is not None and vec2 is not None:
        response = requests.post(server_url, json={'vec1': vec1, 'vec2': vec2})
        if response.status_code == 200:
            data = response.json()
            return data.get('similarity')
    return None

def analyze_modified_network(word_list):
    G = nx.Graph()

    # 添加节点
    for word in word_list:
        G.add_node(word)

    # 首先添加第一个词与所有其他词之间的边
    first_word = word_list[0]
    total_similarity = 0
    count = 0

    for i in range(1, len(word_list)):
        try:
            similarity = calculate_similarity(first_word, word_list[i])
            G.add_edge(first_word, word_list[i], weight=similarity)
            count += 1
            total_similarity += similarity
            if count == 10:
                average_similarity = total_similarity / count

        except KeyError:
            # 如果词不在模型词汇表中
            continue



    # 添加其他词之间的边，条件是语义距离不大于平均距离
    for i in range(1, len(word_list)):
        for j in range(i + 1, len(word_list)):
            try:
                similarity = calculate_similarity(word_list[i], word_list[j])
                if similarity <= average_similarity and similarity>0:
                    G.add_edge(word_list[i], word_list[j], weight=similarity)
            except KeyError:
                continue

    # 检查图是否连通
    if nx.is_connected(G):
        path_leng = nx.average_shortest_path_length(G, weight='weight')
        clustering_coefficient = nx.average_clustering(G, weight='weight')
        return path_leng, clustering_coefficient
    else:
        return  None, None  # 图不连通时返回None
# 启动实验前的准备
def find_matching_result(data_list, final_list, tasks, task_index, path_org, cluster_org):
    for answer in data_list:
        new_list = final_list.copy()  # 创建 final_list 的副本
        new_list.append(answer)  # 在副本上添加新词
        try:
            path_new, cluster_new = analyze_modified_network(new_list)
            # 计算新的路径和簇的差异
            diff_path = path_org - path_new
            diff_cluster = cluster_org - cluster_new
            print(diff_path, diff_cluster)
        except Exception as e:
            print(f"Error processing answer {answer}: {e}")
            continue  # 如果出现异常，继续处理下一个 answer

        # 检查条件
        if tasks[task_index][3] > 0:
            if tasks[task_index][4] > 0 and diff_path > 0 and diff_cluster > 0:
                return tasks[task_index][0], answer, diff_path, diff_cluster
            elif tasks[task_index][4] <0 and diff_path > 0 and diff_cluster < 0:
                return tasks[task_index][0], answer, diff_path, diff_cluster
        else:
            if tasks[task_index][4] > 0 and diff_path < 0 and diff_cluster > 0:
                return tasks[task_index][0], answer, diff_path, diff_cluster
            elif tasks[task_index][4] < 0 and diff_path < 0 and diff_cluster < 0:
                return tasks[task_index][0], answer, diff_path, diff_cluster

    random_answer = random.choice(data_list)
    new_list = final_list.copy()  # 创建 final_list 的副本
    new_list.append(random_answer)  # 在副本上添加新词
    try:
        path_new, cluster_new = analyze_modified_network(new_list)
        # 计算新的路径和簇的差异
        diff_path = path_org - path_new
        diff_cluster = cluster_org - cluster_new
        print(diff_path, diff_cluster)
    except Exception as e:
        print(f"Error processing answer {random_answer}: {e}")

    return tasks[task_index][0], random_answer, diff_path, diff_cluster  # 如果没有找到符合条件的结果，返回None

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

def get_current_time_numeric():
    now = datetime.now()
    formatted_time = now.strftime("%y%m%d%H")
    return formatted_time

start_exp()
numeric_time = get_current_time_numeric()
# 启动事件循环
root.mainloop()
