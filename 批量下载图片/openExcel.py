import os  # 这里导入了 os 模块
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import openpyxl

# 引入之前提取的函数
from original_code import process_excel_file  # 假设原代码保存在 original_code.py


def select_file():
    """打开文件选择对话框，选择 Excel 文件"""
    file_path = filedialog.askopenfilename(
        title="请选择 Excel 文件",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)
        load_sheet_names(file_path)


def load_sheet_names(file_path):
    """加载 Excel 文件中的所有 Sheet 名称到下拉菜单"""
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet_names = wb.sheetnames
        sheet_menu["menu"].delete(0, "end")
        for name in sheet_names:
            sheet_menu["menu"].add_command(label=name, command=tk._setit(sheet_var, name))
        sheet_var.set(sheet_names[0])  # 默认选择第一个 Sheet
    except Exception as e:
        messagebox.showerror("错误", f"加载 Sheet 名称失败：{str(e)}")


def select_output_dir():
    """选择图片保存目录"""
    dir_path = filedialog.askdirectory(title="请选择图片保存目录")
    if dir_path:
        entry_output_dir.delete(0, tk.END)
        entry_output_dir.insert(0, dir_path)


def start_processing():
    """启动文件处理"""
    file_path = entry_file_path.get()
    output_dir = entry_output_dir.get()
    sheet_name = sheet_var.get()
    target_name = entry_name.get().strip()

    if not os.path.isfile(file_path):
        messagebox.showerror("错误", "文件路径无效，请重新选择")
        return

    if not os.path.isdir(output_dir):
        messagebox.showerror("错误", "图片保存路径无效，请重新选择")
        return

    try:
        # 创建进度条
        progressbar = ttk.Progressbar(root, length=300, mode='determinate', maximum=100)
        progressbar.pack(pady=10)

        def update_progress(progress):
            progressbar['value'] = progress
            root.update_idletasks()  # 强制更新UI

        # 调用原代码的处理函数，传递进度回调
        process_excel_file(file_path, output_dir, sheet_name, target_name, update_progress)
        messagebox.showinfo("成功", "文件处理完成！")
    except Exception as e:
        messagebox.showerror("错误", f"文件处理出错：\n{str(e)}")


# 创建主窗口
root = tk.Tk()
root.title("Excel 图片导入处理")

# 文件路径选择
frame = tk.Frame(root)
frame.pack(padx=20, pady=10)

label_file_path = tk.Label(frame, text="文件路径:")
label_file_path.grid(row=0, column=0, padx=5, pady=5)

entry_file_path = tk.Entry(frame, width=50)
entry_file_path.grid(row=0, column=1, padx=5, pady=5)

btn_browse_file = tk.Button(frame, text="浏览", command=select_file)
btn_browse_file.grid(row=0, column=2, padx=5, pady=5)

# 输出目录选择
label_output_dir = tk.Label(frame, text="保存目录:")
label_output_dir.grid(row=1, column=0, padx=5, pady=5)

entry_output_dir = tk.Entry(frame, width=50)
entry_output_dir.grid(row=1, column=1, padx=5, pady=5)

btn_browse_dir = tk.Button(frame, text="选择", command=select_output_dir)
btn_browse_dir.grid(row=1, column=2, padx=5, pady=5)

# Sheet 名称选择
label_sheet = tk.Label(frame, text="选择 Sheet:")
label_sheet.grid(row=2, column=0, padx=5, pady=5)

sheet_var = tk.StringVar(root)
sheet_menu = tk.OptionMenu(frame, sheet_var, [])
sheet_menu.grid(row=2, column=1, padx=5, pady=5, sticky="w")

# 姓名输入
label_name = tk.Label(frame, text="输入姓名:")
label_name.grid(row=3, column=0, padx=5, pady=5)

entry_name = tk.Entry(frame, width=50)
entry_name.grid(row=3, column=1, padx=5, pady=5)

# 启动按钮
btn_process = tk.Button(root, text="开始处理", command=start_processing)
btn_process.pack(pady=10)

# 主循环
root.mainloop()
