import os
import sys
import cx_Oracle
import pandas as pd
import configparser
from cryptography.fernet import Fernet

# 加载加密密钥
def load_key():
    try:
        key = open('secret.key', 'rb').read()
        return key
    except Exception as e:
        print(f"Failed to load encryption key: {e}")
        sys.exit(1)

# 解密密码
def decrypt_password(encrypted_password):
    try:
        key = load_key()
        cipher_suite = Fernet(key)
        decoded_password = cipher_suite.decrypt(encrypted_password.encode()).decode()
        return decoded_password
    except Exception as e:
        print(f"Failed to decrypt password: {e}")
        sys.exit(1)

# 读取配置文件
def read_config(config_file='config.ini'):
    config = configparser.ConfigParser()
    config.read(config_file)
    return config

# 读取 SQL 文件并提取 SQL 语句和对应的 Sheet 名字
def read_sql_queries(sql_file):
    try:
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
        return queries
    except Exception as e:
        print(f"Failed to read SQL file: {e}")
        sys.exit(1)

# 执行 SQL 查询并将结果保存到 Excel
def export_to_excel(sql_file, output_file):
    config = read_config()
    user = config['database']['user']
    encrypted_password = config['database']['password']
    password = decrypt_password(encrypted_password)
    dsn = config['database']['dsn']

    queries = read_sql_queries(sql_file)

    try:
        connection = cx_Oracle.connect(user=user, password=password, dsn=dsn)
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for i, (sheet_name, query) in enumerate(queries):
                df = pd.read_sql(query, con=connection)
                sheet_name = sheet_name if sheet_name else f"Sheet{i+1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Data exported successfully to {output_file}")
    except Exception as e:
        print(f"Failed to export data: {e}")
        sys.exit(1)
    finally:
        if 'connection' in locals():
            connection.close()

# 主函数
def main():
    if len(sys.argv) != 3:
        print("Usage: python oracle_report_batch.py <sql_file> <output_file>")
        sys.exit(1)

    sql_file = sys.argv[1]
    output_file = sys.argv[2]

    if not os.path.exists(sql_file):
        print(f"SQL file does not exist: {sql_file}")
        sys.exit(1)

    export_to_excel(sql_file, output_file)

if __name__ == "__main__":
    main()
