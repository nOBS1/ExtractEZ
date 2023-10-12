import pandas as pd
import tkinter as tk
from tkinter import filedialog

# 定义全局变量存储文件名和筛选信息
input_file = ""
filter_text = ""

# 定义函数来读取Excel文件
def choose_file():
    global input_file
    input_file = filedialog.askopenfilename(title="1. 选择Excel文件")
    if input_file:
        file_label.config(text=f"已打开文件: {input_file}")

# 定义函数来筛选并保存Excel文件
def filter_and_save():
    global input_file, filter_text
    if input_file:
        try:
            # 读取用户选择的Excel文件
            df = pd.read_excel(input_file)
            
            # 获取用户输入的筛选信息
            filter_text = filter_entry.get()
            
            if filter_text:
                # 使用 str.contains 方法筛选处理人列
                filtered_df = df[df['处理人'].str.contains(filter_text)]
                
                # 选择保存位置
                output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
                if output_file:
                    # 保存筛选后的记录到新的Excel文件
                    filtered_df.to_excel(output_file, index=False)
                    result_label.config(text=f"筛选后的记录已保存到 {output_file}")
                else:
                    result_label.config(text="未选择保存位置")
            else:
                result_label.config(text="请输入筛选信息")
        except Exception as e:
            result_label.config(text=f"出现错误：{str(e)}")
    else:
        file_label.config(text="请选择Excel文件")

# 创建GUI窗口
root = tk.Tk()
root.title("Excel筛选工具")

# 添加选择文件按钮
file_button = tk.Button(root, text="1. 选择Excel文件", command=choose_file)
file_button.pack()

# 添加标签来显示已打开的文件
file_label = tk.Label(root, text="", wraplength=400)
file_label.pack()

# 添加输入框来获取筛选信息
filter_entry = tk.Entry(root, width=30)
filter_entry.pack()

# 添加结果标签
result_label = tk.Label(root, text="", wraplength=400)
result_label.pack()

# 添加保存按钮
save_button = tk.Button(root, text="2. 保存筛选结果", command=filter_and_save)
save_button.pack()

# 运行GUI程序
root.mainloop()
