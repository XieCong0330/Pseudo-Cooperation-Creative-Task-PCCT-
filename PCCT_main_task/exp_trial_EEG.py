import tkinter as tk
from tkinter import messagebox
import random
import pandas as pd
import networkx as nx
import requests
import time
from datetime import datetime
import os
import copy
import ctypes
pp = 0x378
inpout = ctypes.windll.inpoutx64
# 准备词汇列表，每个词只列一次
def write_pa(value):
    inpout.Out32(pp,value)

def send_mark(value):
    write_pa(value)
    time.sleep(0.001)
    write_pa(0)
    print(value)


base_words = ["砖头","便利贴","筷子","钥匙","报纸"]
diss_words = ["沙漠","宇宙","钢琴","袜子","西瓜"]
control_words = ["管尼","拉明","格路","帕本","范谢"]
oct_words = pd.read_excel('oct.xls',header=None)
random.shuffle(base_words)
# 先生成所有词的联想任务
# Generating tasks as dictionaries
control_tasks = [[word, "请重复该文字内容", 10000] for word in control_words for _ in range(2)]
oct_tasks = [[word, "请描述特征", 20000] for word in oct_words[0]]
oct_tasks = [sublist + [oct_words[1][i]] for i, sublist in enumerate(oct_tasks)]
association_tasks = [[word, "联想", 10000] for word in base_words for _ in range(10)]
disassociation_tasks = [[word, "分离", 10000] for word in diss_words for _ in range(10)]
association_tasks = association_tasks+control_tasks
disassociation_tasks = disassociation_tasks+control_tasks
pattern = [(1, 1), (1, 1), (1, -1), (1, -1), (-1, 1), (-1, 1), (-1, -1), (-1, -1)]

# Generating creativity tasks with varying last two elements based on the specified pattern
creativity_tasks = []
for word in base_words:
    random.shuffle(pattern)
    for i in range(8):  # Create 8 entries per word
        values = pattern[i]  # Select values from the pattern based on index
        creativity_tasks.append([word, "创造性", 20000] + list(values))

org_list = []
timers = {}
# Randomize the tasks
random.shuffle(association_tasks)
random.shuffle(disassociation_tasks)
random.shuffle(creativity_tasks)
random.shuffle(oct_tasks)

tasks = association_tasks+disassociation_tasks+creativity_tasks+oct_tasks

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
submit_sele_button = tk.Button(root, text="提交他人观点", command=lambda: submit_sele())
submit_sele_button.pack(pady=20)
continuedis_button = tk.Button(root, text="继续", command=lambda: start_dis())
continuedis_button.pack(pady=20)
continueAUT_button = tk.Button(root, text="继续", command=lambda: startAUT_exp())
continueAUT_button.pack(pady=20)
continueoct_button = tk.Button(root, text="继续", command=lambda: startoct_exp())
continueoct_button.pack(pady=20)

task_index = 0
association_list = []
disassociation_list =[]
AUT_list = []
example_list =[]
accpt_idea =[]
# 提交联想或创造性用途
rating_var = tk.IntVar()  # This variable will store the user's rating
rating_frame = tk.Frame(root)
tk.Label(rating_frame, text="1为非常不具有创造性，5为非常有创造性:", font=("Times",30)).pack()
result_list = []
# Create and pack radio buttons within the frame
for i in range(1, 6):
    tk.Radiobutton(rating_frame, text=str(i), variable=rating_var, value=i, font=("Times", 30)).pack(side=tk.LEFT)

def start_exp():
    entry.pack_forget()
    submit_button.pack_forget()
    continueAUT_button.pack_forget()
    continuedis_button.pack_forget()
    continueoct_button.pack_forget()
    continue_button.pack()
    rate_button.pack_forget()
    sele_button.pack_forget()
    submit_sele_button.pack_forget()
    display_label.config(text=f"您好，欢迎来到本实验，本实验分为四个部分，第一部分为自由联想任务，"
                              f"需要您看见提示词后，以最快的速度联想该提示词的一个相关的词语，如：夏天联想到西瓜，"
                              f"联想结束后请按下空格键，并输入您的联想内容。注意，请不要联想特定的名字(如：李明)"
                              f"准备好后请点击‘继续’",font=("Times",30))  # 显示当前任务
    send_mark(1)


def start_dis():
    entry.pack_forget()
    submit_button.pack_forget()
    continueAUT_button.pack_forget()
    continuedis_button.pack_forget()
    continueoct_button.pack_forget()
    continue_button.pack()
    rate_button.pack_forget()
    sele_button.pack_forget()
    submit_sele_button.pack_forget()
    display_label.config(text=f"您好，欢迎来到本实验，本实验分为四个部分，第二部分为分离联想任务，"
                              f"需要您看见提示词后，以最快的速度联想该提示词的一个无关的词语，如：夏天与宇宙无关，"
                              f"联想结束后请按下空格键，并输入您的联想内容。注意，请不要联想特定的名字(如：李明)"
                              f"准备好后请点击‘继续’",font=("Times",30))  # 显示当前任务
    send_mark(2)

def startAUT_exp():
    entry.pack_forget()
    submit_button.pack_forget()
    continueAUT_button.pack_forget()
    continueoct_button.pack_forget()
    continue_button.pack()
    rate_button.pack_forget()
    sele_button.pack_forget()
    submit_sele_button.pack_forget()
    display_label.config(text=f"您好，接下来为本实验的第三部分，第三部分为多用途任务，"
                              f"在该部分，您需要与合作者一起思考某个物品的创造性用途，如：小刀的用途是开锁，"
                              f"您所在的实验组为第二组，因此，您会首先得到合作者的答案，再进行思考，思考结束后请按下空格键，并输入您的答案。"
                              f"准备好后请点击‘继续’。",font=("Times",30))  # 显示当前任务
    send_mark(3)
# 准备每个试次
def startoct_exp():
    entry.pack_forget()
    submit_button.pack_forget()
    continueAUT_button.pack_forget()
    continueoct_button.pack_forget()
    continue_button.pack()
    rate_button.pack_forget()
    sele_button.pack_forget()
    submit_sele_button.pack_forget()
    display_label.config(text=f"您好，接下来为本实验的最后一个部分，为物品描述任务，"
                              f"在该部分，您需要与合作者一起思考某个物品的特征，如：小刀的特征是锋利的，"
                              f"您所在的实验组为第二组，因此，您会首先得到合作者的答案，再进行思考，思考结束后请按下空格键，并输入您的答案。"
                              f"准备好后请点击‘继续’。",font=("Times",30))  # 显示当前任务
    send_mark(4)


def prepare_trial():
    global result
    entry.pack_forget()
    submit_button.pack_forget()
    continuedis_button.pack_forget()
    continue_button.pack_forget()
    rate_button.pack_forget()
    sele_button.pack_forget()
    submit_sele_button.pack_forget()
    result = []
    if task_index < len(tasks):
       display_label.config(text="+", fg="black",font=("Times",40,"bold"))  # 显示注视点
       send_mark(10)
    # 设置延迟后显示任务
       if tasks[task_index][1] == '创造性' or tasks[task_index][1] == '请描述特征':
           delay = random.randint(5000, 10000)
           root.after(delay, lambda: show_example(tasks[task_index]))
       else:
           delay = random.randint(1000, 3000)
           root.after(delay, lambda: show_task(tasks[task_index]))
    else:
        finish_experiment()

def show_example(task):
    entry.pack_forget()
    submit_button.pack_forget()
    continue_button.pack_forget()
    rating_frame.pack_forget()
    if tasks[task_index][1] == '创造性':
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
        similarity_matrix = generate_similarity_matrix(final_list)
        path_org, cluster_org = analyze_modified_network(similarity_matrix, final_list)
        print(path_org, cluster_org)
        data = pd.read_excel('new_AUT3.xls', sheet_name=tasks[task_index][0])

        # 将DataFrame转换为列表
        # 如果你只想转换某个特定列，可以使用 data['column_name'].tolist()
        data = data.iloc[:, 0]
        data_list = data.values.tolist()
        first_column = [row[0] for row in result_list]
        data_list = [item for item in data_list if item not in first_column]
        random.shuffle(data_list)
        result_exp = find_matching_result(data_list, final_list, tasks, task_index, path_org, cluster_org,similarity_matrix)
        result.append(result_exp[1])
        result.append(result_exp[2])
        result.append(result_exp[3])
        example_list.append(result_exp)
        example = result_exp[1]
        print(example)
        send_mark(21)
        display_label.config(text=f"{task[0]}的用途是： {example} ", fg="black",font=("Times",30))  # 显示当前任务
    else:
        send_mark(22)
        display_label.config(text=f"{task[0]}的特征是： {task[3]} ", fg="black", font=("Times", 30))  # 显示当前任务
    root.after(5000, lambda: show_task(tasks[task_index]))
# 显示当前任务
def show_task(task):
    global start_time
    continue_button.pack_forget()
    send_mark(11)
    display_label.config(text=f" {task[1]} '{task[0]}'，思考结束后按空格", fg="black",font=("Times",30))  # 显示当前任务
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
    send_mark(12)
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
    sim = calculate_similarity(tasks[task_index][0], user_input) # 立刻计算语义距离不用就删掉
    if sim == None  and tasks[task_index][1] != '请重复该文字内容':
        send_mark(66)
        messagebox.showinfo("无法理解", "无法理解该答案，请重新作答。")  # 显示提示信息
        entry.pack_forget()
        submit_button.pack_forget()
        show_task(tasks[task_index])
        return
    if tasks[task_index][1] == '联想':
        send_mark(31)
        judge_len = [item for item in association_list if item[0].strip() == tasks[task_index][0]]
        words = [item[1] for item in judge_len]
        if user_input in words or user_input == tasks[task_index][0]:
            send_mark(66)
            messagebox.showinfo("无效答案", "该答案重复，请重新作答。")  # 显示提示信息
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
            similarity_matrix = generate_similarity_matrix(words)
            path_len, clustering_coefficient = analyze_modified_network(similarity_matrix,words) #立刻计算语义网络的拓扑属性，如果只收数据就不用
            org = [tasks[task_index][0], path_len, clustering_coefficient]
            org_list.append(org)
            result.append(path_len)
            result.append(clustering_coefficient)
            result_list.append(result)

            print(org_list)
        else:
            result_list.append(result)
    elif tasks[task_index][1] == '分离':
        send_mark(32)
        judge_len = [item for item in disassociation_list if item[0].strip() == tasks[task_index][0]]
        words = [item[1] for item in judge_len]
        if user_input in words or user_input == tasks[task_index][0]:
            send_mark(66)
            messagebox.showinfo("无效答案", "该答案重复，请重新作答。")  # 显示提示信息
            entry.pack_forget()
            submit_button.pack_forget()
            show_task(tasks[task_index])
            return
        else:
            disassociation_list.append([tasks[task_index][0], user_input, sim])
            judge_len = [item for item in disassociation_list if item[0].strip() == tasks[task_index][0]]
            words = [item[1] for item in judge_len]
        print(words)
        result_list.append(result)

    elif tasks[task_index][1] == '创造性':
        send_mark(33)
        AUT_list.append([tasks[task_index][0], user_input, sim])

    elif tasks[task_index][1] == '请重复该文字内容':
        send_mark(34)
        if user_input != tasks[task_index][0]:
            send_mark(66)
            messagebox.showinfo("无效答案", "请重复输入词语")  # 显示提示信息
            entry.pack_forget()
            submit_button.pack_forget()
            show_task(tasks[task_index])
            return


    task_index += 1
    if task_index == len(association_tasks):
        show_rest_screen()
    elif task_index == len(association_tasks)+len(disassociation_tasks):
        show_rest_screen()
    elif tasks[task_index - 1][1] == '创造性':
        show_rating_screen()
    elif task_index >= len(tasks):
        finish_experiment()
    else:
        prepare_trial()





# 结束实验
def show_rating_screen():
    send_mark(23)
    entry.pack_forget()
    submit_button.pack_forget()
    rating_var.set(0)

    display_label.config(text=f"您的答案是：{entry.get()}，请问您认为该答案的创造性水平为何", fg="black",font=("Times",30))
    rating_frame.pack(pady=20)  # Show the rating frame for creativity level

    rate_button.pack(pady=20)

def show_examp_rate():
    send_mark(24)
    entry.pack_forget()
    submit_button.pack_forget()
    rate_button.pack_forget()
    self_rate = rating_var.get()
    result.append(self_rate)
    rating_var.set(0)
    example = example_list[task_index-len(association_tasks)-len(disassociation_tasks)-1][1]

    display_label.config(text=f"他人的答案是：{example}，请问您认为该答案的创造性水平为何", fg="black",font=("Times",30))

    sele_button.pack(pady=20)

def show_examp_sele():
    send_mark(25)
    root.unbind('<space>')
    entry.delete(0, tk.END)
    sele_button.pack_forget()
    rating_frame.pack_forget()
    self_rate = rating_var.get()
    result.append(self_rate)
    display_label.config(text=f"请问您的答案是否和他人的答案相关联，如果是，请输入他人的答案，如果否，请填写‘没有相关答案’", fg="black",font=("Times",30))
    entry.pack(pady=20)
    entry.focus_set()
    submit_sele_button.pack(pady=20)

def submit_sele():
    send_mark(26)
    input_idea = entry.get()
    judge_len = [item for item in example_list if item[0].strip() == tasks[task_index-1][0]]
    result.append(input_idea)
    result_list.append(result)

    for item in judge_len:
        sim = calculate_similarity(input_idea, item[1])
        if sim < 0.2:
            accpt_idea.append(item)
    entry.delete(0, tk.END)
    if task_index == len(creativity_tasks) + len(association_tasks) + len(disassociation_tasks):
        show_rest_screen()
    else:
        prepare_trial()


def show_rest_screen():
    send_mark(77)
    entry.pack_forget()
    submit_button.pack_forget()
    submit_sele_button.pack_forget()
    if task_index == len(association_tasks):
        continuedis_button.pack()
    elif task_index == len(disassociation_tasks)+len(association_tasks):
        continueAUT_button.pack()
    else:
        continueoct_button.pack()
    display_label.config(text="休息一下, 按 '继续' 结束休息", fg="black",font=("Times",30,"bold"))

def finish_experiment():
    send_mark(88)
    display_label.config(text="实验结束，请联系主试", fg="black",font=("Times",30,"bold"))
    entry.pack_forget()
    submit_button.pack_forget()
    continue_button.pack_forget()

    save_to_excel(result_list, filename=f'{numeric_time}.xlsx',sheet_name='data')

# 计算语义距离与语义网络，如果不用全部删了即可

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
def generate_similarity_matrix(word_list):
    n = len(word_list)
    similarity_matrix = [[0.0 for _ in range(n)] for _ in range(n)]  # 初始化连接矩阵

    for i in range(n):
        for j in range(i, n):  # 只需计算上三角部分，因为矩阵是对称的
            if i == j:
                similarity_matrix[i][j] = 1.0  # 同一个词的相似度为 1
            else:
                try:
                    similarity = calculate_similarity(word_list[i], word_list[j])
                    similarity_matrix[i][j] = similarity
                    similarity_matrix[j][i] = similarity  # 对称位置赋值
                except KeyError:
                    # 如果词不在模型词汇表中，设置为 None 或 0
                    similarity_matrix[i][j] = None
                    similarity_matrix[j][i] = None

    return similarity_matrix


def update_similarity_matrix(word_list, new_word, similarity_matrix):
    # 新词与旧词的相似度计算
    n = len(word_list)
    sm = copy.deepcopy(similarity_matrix)  # 使用深拷贝
    # 为新词添加一行并初始化该行（所有值为 None）
    sm.append([None] * (n + 1))  # 新词的行，初始所有值为 None

    # 更新每个旧词与新词之间的相似度
    for i in range(n):
        try:
            # 计算新词与旧词的相似度
            similarity = calculate_similarity(word_list[i], new_word)
            sm[i].append(similarity)  # 更新旧词与新词的相似度
            sm[n][i] = similarity  # 更新新词与旧词的相似度
        except KeyError:
            # 如果某个词不在模型词汇表中，设置为 None
            sm[i].append(None)
            sm[n][i] = None

    # 新词与自己相似度为 1
    sm[n].append(1.0)  # 新词与自己相似度为 1.0

    return sm


def analyze_modified_network(similarity_matrix, word_list):
    G = nx.Graph()

    # 添加节点
    for word in word_list:
        G.add_node(word)

    # 首先添加第一个词与所有其他词之间的边
    first_word = word_list[0]
    total_similarity = 0
    count = 0

    for i in range(1, len(word_list)):
        similarity = similarity_matrix[0][i]  # 获取第一个词与第i个词的语义距离
        if similarity is not None:  # 确保语义距离有效
            G.add_edge(first_word, word_list[i], weight=similarity)
            count += 1
            total_similarity += similarity
            if count == 10:
                average_similarity = total_similarity / count

    # 添加其他词之间的边，条件是语义距离不大于平均距离
    for i in range(1, len(word_list)):
        for j in range(i + 1, len(word_list)):
            similarity = similarity_matrix[i][j]  # 获取第i个词与第j个词的语义距离
            if similarity is not None and similarity <= average_similarity and similarity > 0:
                G.add_edge(word_list[i], word_list[j], weight=similarity)

    # 检查图是否连通
    if nx.is_connected(G):
        path_length = nx.average_shortest_path_length(G, weight='weight')
        clustering_coefficient = nx.average_clustering(G, weight='weight')
        return path_length, clustering_coefficient
    else:
        return None, None  # 图不连通时返回None
# 启动实验前的准备
def find_matching_result(data_list, final_list, tasks, task_index, path_org, cluster_org,similarity_matrix):
    for answer in data_list:
        new_list = final_list.copy()  # 创建 final_list 的副本
        new_list.append(answer)  # 在副本上添加新词
        update_matrix = update_similarity_matrix(final_list, answer, similarity_matrix)
        try:
            path_new, cluster_new = analyze_modified_network(update_matrix,new_list)
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
