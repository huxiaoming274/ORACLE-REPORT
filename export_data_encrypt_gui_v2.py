import os
import sys
import pandas as pd
import configparser
import tkinter as tk
from tkinter import filedialog, messagebox
from cryptography.fernet import Fernet

# 用 python-oracledb（默认薄模式）
import oracledb as db

# 加载加密密钥
def load_key():
    try:
        with open('secret.key', 'rb') as f:
            return f.read()
    except Exception as e:
        messagebox.showerror("错误 / Error", f"加载密钥失败: {e} / Failed to load key: {e}")
        raise

# 解密密码
def decrypt_password(encrypted_password):
    try:
        key = load_key()
        from cryptography.fernet import Fernet
        cipher_suite = Fernet(key)
        return cipher_suite.decrypt(encrypted_password.encode()).decode()
    except Exception as e:
        messagebox.showerror("错误 / Error", f"密码解密失败: {e} / Password decryption failed: {e}")
        raise

def read_config(config_file):
    config = configparser.ConfigParser()
    config.read(config_file, encoding="utf-8")
    return config

# 兼容 host:port:SID => host:port/SID
def normalize_dsn(raw_dsn: str) -> str:
    dsn = raw_dsn.strip()
    if dsn.count(':') == 2 and '/' not in dsn:
        host, port, sid = dsn.split(':', 2)
        dsn = f"{host}:{port}/{sid}"
    return dsn

def test_connection(config_file):
    config = read_config(config_file)
    user = config['database']['user']
    encrypted_password = config['database']['password']
    try:
        password = decrypt_password(encrypted_password)
    except Exception:
        return
    dsn = normalize_dsn(config['database']['dsn'])

    try:
        # 薄模式连接（不要调用 init_oracle_client，不要 SYSDBA）
        conn = db.connect(user=user, password=password, dsn=dsn)
        conn.close()
        messagebox.showinfo("成功 / Success", "数据库连接成功 / Database connection successful")
    except Exception as e:
        messagebox.showerror("错误 / Error", f"数据库连接失败: {e} / Database connection failed: {e}")

def select_sql_file():
    file_path = filedialog.askopenfilename(title="选择SQL文件 / Select SQL File",
                                           filetypes=[("SQL Files", "*.sql")])
    sql_file_var.set(file_path)

def select_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", title="保存为 / Save As",
                                             filetypes=[("Excel Files", "*.xlsx")])
    output_file_var.set(file_path)

def run_export():
    config_file = 'config.ini'
    sql_file = sql_file_var.get()
    output_file = output_file_var.get()
    if not sql_file or not output_file:
        messagebox.showwarning("警告 / Warning", "请填写所有字段 / Please fill in all fields")
        return
    export_data_to_excel(config_file, sql_file, output_file)

def load_db_info():
    config = read_config('config.ini')
    db_info.set(f"User: {config['database']['user']}\nDSN: {config['database']['dsn']}")

def export_data_to_excel(config_file, sql_file, output_file):
    config = read_config(config_file)
    user = config['database']['user']
    encrypted_password = config['database']['password']
    password = decrypt_password(encrypted_password)
    dsn = normalize_dsn(config['database']['dsn'])

    queries = read_sql_queries(sql_file)
    conn = db.connect(user=user, password=password, dsn=dsn)

    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for i, (sheet_name, query) in enumerate(queries):
                df = pd.read_sql(query, con=conn)
                if not sheet_name:
                    sheet_name = f"Sheet{i+1}"
                sheet_name = sheet_name[:31]  # Excel 限制
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        messagebox.showinfo("成功 / Success",
                            f"数据已成功导出到 {output_file} / Data has been successfully exported to {output_file}")
    except Exception as e:
        messagebox.showerror("错误 / Error", f"导出数据时出错: {e} / Error exporting data: {e}")
    finally:
        conn.close()

def read_sql_queries(sql_file):
    with open(sql_file, 'r', encoding='utf-8') as f:
        content = f.read()
    queries = []
    for part in content.split(';'):
        lines = part.strip().split('\n')
        if not lines:
            continue
        sheet_name = None
        if lines[0].strip().lower().startswith('-- sheet_name:'):
            sheet_name = lines[0].split(':', 1)[1].strip()
            lines = lines[1:]
        query = '\n'.join(lines).strip()
        if query:
            queries.append((sheet_name, query))
    return queries

# GUI
root = tk.Tk()
root.title("Oracle数据导出工具 / Oracle Data Export Tool")

tk.Label(root, text="选择SQL文件 / Select SQL File:").grid(row=0, column=0, padx=10, pady=5)
sql_file_var = tk.StringVar()
tk.Entry(root, textvariable=sql_file_var, width=50).grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="浏览 / Browse", command=select_sql_file).grid(row=0, column=2, padx=10, pady=5)

tk.Label(root, text="导出文件名 / Output File Name:").grid(row=1, column=0, padx=10, pady=5)
output_file_var = tk.StringVar()
tk.Entry(root, textvariable=output_file_var, width=50).grid(row=1, column=1, padx=10, pady=5)
tk.Button(root, text="保存为 / Save As", command=select_output_file).grid(row=1, column=2, padx=10, pady=5)

tk.Button(root, text="执行 / Execute", command=run_export).grid(row=2, column=0, columnspan=3, pady=20)

db_info = tk.StringVar()
tk.Label(root, text="数据库配置信息 / Database Configuration Info:").grid(row=3, column=0, padx=10, pady=5)
tk.Message(root, textvariable=db_info, width=400).grid(row=3, column=1, padx=10, pady=5)
tk.Button(root, text="加载配置信息 / Load Configuration Info", command=load_db_info).grid(row=3, column=2, padx=10, pady=5)

tk.Button(root, text="测试连接 / Test Connection",
          command=lambda: test_connection('config.ini')).grid(row=4, column=0, columnspan=3, pady=20)

load_db_info()
root.mainloop()
