import os
import tkinter as tk
from tkinter import filedialog, messagebox
from cryptography.fernet import Fernet

# 加密密钥文件路径
KEY_FILE = 'secret.key'

def generate_key():
    key = Fernet.generate_key()
    with open(KEY_FILE, 'wb') as key_file:
        key_file.write(key)
    return key

def load_key():
    if not os.path.exists(KEY_FILE):
        return generate_key()
    return open(KEY_FILE, 'rb').read()

def encrypt_string(plain_text):
    key = load_key()
    cipher_suite = Fernet(key)
    cipher_text = cipher_suite.encrypt(plain_text.encode())
    return cipher_text.decode()

def save_key():
    key = generate_key()
    messagebox.showinfo("密钥已生成 / Key Generated", f"加密密钥已生成并保存在 {KEY_FILE} / Encryption key has been generated and saved in {KEY_FILE}")

def encrypt_text():
    plain_text = input_text.get("1.0", tk.END).strip()
    if not plain_text:
        messagebox.showwarning("警告 / Warning", "请输入要加密的字符串 / Please enter the string to encrypt")
        return

    encrypted_text = encrypt_string(plain_text)
    output_text.delete("1.0", tk.END)
    output_text.insert(tk.END, encrypted_text)
    messagebox.showinfo("加密成功 / Encryption Successful", "字符串已成功加密 / The string has been successfully encrypted")

# 创建GUI
root = tk.Tk()
root.title("字符串加密工具 / String Encryption Tool")

tk.Label(root, text="输入要加密的字符串 / Enter the string to encrypt:").grid(row=0, column=0, padx=10, pady=5)
input_text = tk.Text(root, height=5, width=50)
input_text.grid(row=1, column=0, padx=10, pady=5)

tk.Button(root, text="加密 / Encrypt", command=encrypt_text).grid(row=2, column=0, pady=10)

tk.Label(root, text="加密后的字符串 / Encrypted string:").grid(row=3, column=0, padx=10, pady=5)
output_text = tk.Text(root, height=5, width=50)
output_text.grid(row=4, column=0, padx=10, pady=5)

tk.Button(root, text="生成新的加密密钥 / Generate New Key", command=save_key).grid(row=5, column=0, pady=10)

root.mainloop()
