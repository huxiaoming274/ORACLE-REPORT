import os
import sys
import cx_Oracle
import pandas as pd
import configparser
import tkinter as tk
from tkinter import filedialog, messagebox

# 设置 Oracle Instant Client 库路径
instant_client_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'instantclient_19_9')
if sys.platform == "win32":
    os.environ["PATH"] = instant_client_dir + ";" + os.environ["PATH"]
else:
    os.environ["LD_LIBRARY_PATH"] = instant_client_dir + ":" + os.environ.get("LD_LIBRARY_PATH", "")

def read_config(config_file):
    config = configparser.ConfigParser()
    config.read(config_file)
    return config

def read_sql_queries(sql_file):
    with open(sql_file, 'r') as file:
        content = file.read()

    queries = []
    parts = content.split(';')
    for part in parts:
        lines = part.strip().split('\n')
        if lines:
            sheet_name = None
            if lines[0].strip().startswith('-- sheet_name:'):
                sheet_name = lines[0].strip().split(':', 1)[1].strip()
                lines = lines[1:]  # 去掉注释行
            query = '\n'.join(lines).strip()
            if query:
                queries.append((sheet_name, query))
                print(f"Read query: {query} with sheet name: {sheet_name}")  # 调试信息
    return queries

def export_data_to_excel(config_file, sql_file, output_file):
    config = read_config(config_file)
    user = config['database']['user']
    password = config['database']['password']
    dsn = config['database']['dsn']

    queries = read_sql_queries(sql_file)

    connection = cx_Oracle.connect(user=user, password=password, dsn=dsn)

    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for i, (sheet_name, query) in enumerate(queries):
                df = pd.read_sql(query, con=connection)
                if not sheet_name:
                    sheet_name = f'Sheet{i+1}'
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Exported {query} to sheet: {sheet_name}")  # 调试信息

        messagebox.showinfo("成功", f"数据已成功导出到{output_file}")

    except Exception as e:
        messagebox.showerror("错误", f"导出数据时出错: {e}")
        print(f"Error exporting data: {e}")  # 调试信息

    finally:
        connection.close()

def test_connection(config_file):
    config = read_config(config_file)
    user = config['database']['user']
    password = config['database']['password']
    dsn = config['database']['dsn']

    try:
        connection = cx_Oracle.connect(user=user, password=password, dsn=dsn)
        connection.close()
        messagebox.showinfo("成功", "数据库连接成功")
    except Exception as e:
        messagebox.showerror("错误", f"数据库连接失败: {e}")

def select_sql_file():
    file_path = filedialog.askopenfilename(title="选择SQL文件", filetypes=[("SQL Files", "*.sql")])
    sql_file_var.set(file_path)

def select_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", title="保存为", filetypes=[("Excel Files", "*.xlsx")])
    output_file_var.set(file_path)

def run_export():
    config_file = 'config.ini'
    sql_file = sql_file_var.get()
    output_file = output_file_var.get()

    if not sql_file or not output_file:
        messagebox.showwarning("警告", "请填写所有字段")
        return

    export_data_to_excel(config_file, sql_file, output_file)

def load_db_info():
    config = read_config('config.ini')
    db_info.set(f"User: {config['database']['user']}\nPassword: {config['database']['password']}\nDSN: {config['database']['dsn']}")

# 创建GUI
root = tk.Tk()
root.title("Oracle数据导出工具")

tk.Label(root, text="选择SQL文件:").grid(row=0, column=0, padx=10, pady=5)
sql_file_var = tk.StringVar()
tk.Entry(root, textvariable=sql_file_var, width=50).grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="浏览", command=select_sql_file).grid(row=0, column=2, padx=10, pady=5)

tk.Label(root, text="导出文件名:").grid(row=1, column=0, padx=10, pady=5)
output_file_var = tk.StringVar()
tk.Entry(root, textvariable=output_file_var, width=50).grid(row=1, column=1, padx=10, pady=5)
tk.Button(root, text="保存为", command=select_output_file).grid(row=1, column=2, padx=10, pady=5)

tk.Button(root, text="执行", command=run_export).grid(row=2, column=0, columnspan=3, pady=20)

# 添加数据库信息区域
db_info = tk.StringVar()
tk.Label(root, text="数据库配置信息:").grid(row=3, column=0, padx=10, pady=5)
tk.Message(root, textvariable=db_info, width=400).grid(row=3, column=1, padx=10, pady=5)
tk.Button(root, text="加载配置信息", command=load_db_info).grid(row=3, column=2, padx=10, pady=5)

# 添加测试连接按钮
tk.Button(root, text="测试连接", command=lambda: test_connection('config.ini')).grid(row=4, column=0, columnspan=3, pady=20)

# 加载初始数据库信息
load_db_info()

root.mainloop()
