import os
import sys
import oracledb  # 使用 oracledb 的 Thin 模式
import pandas as pd
import configparser
from cryptography.fernet import Fernet
import logging
from datetime import datetime

# 配置 oracledb 使用 Thin 模式
oracledb.init_oracle_client(lib_dir=None)  # Thin 模式下无需指定 lib_dir

# 日志目录和文件名
LOG_DIR = os.path.join(os.getcwd(), "log")
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)
LOG_FILE = os.path.join(LOG_DIR, f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")

# 配置日志
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# 加载加密密钥
def load_key():
    try:
        key = open('secret.key', 'rb').read()
        logging.info("Encryption key loaded successfully.")
        return key
    except Exception as e:
        logging.error(f"Failed to load encryption key: {e}")
        sys.exit(1)

# 解密密码
def decrypt_password(encrypted_password):
    try:
        key = load_key()
        cipher_suite = Fernet(key)
        decoded_password = cipher_suite.decrypt(encrypted_password.encode()).decode()
        logging.info("Password decrypted successfully.")
        return decoded_password
    except Exception as e:
        logging.error(f"Failed to decrypt password: {e}")
        sys.exit(1)

# 读取配置文件
def read_config(config_file='config.ini'):
    config = configparser.ConfigParser()
    try:
        config.read(config_file)
        logging.info(f"Configuration file {config_file} loaded successfully.")
        return config
    except Exception as e:
        logging.error(f"Failed to read configuration file: {e}")
        sys.exit(1)

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
        logging.info(f"SQL file {sql_file} parsed successfully.")
        return queries
    except Exception as e:
        logging.error(f"Failed to read SQL file: {e}")
        sys.exit(1)

# 执行 SQL 查询并将结果保存到 Excel
def export_to_excel(sql_file, output_file):
    """
    Exports data from an Oracle database to an Excel file.
    This function reads SQL queries from a specified file, executes them against an Oracle database,
    and writes the results to an Excel file with multiple sheets.
    Args:
        sql_file (str): Path to the file containing SQL queries. Each query should be associated with a sheet name.
        output_file (str): Path to the output Excel file.
    Raises:
        Exception: If there is an error during database connection, query execution, or file writing.
    Notes:
        - The function reads database configuration from a configuration file.
        - The database password is expected to be encrypted and will be decrypted before use.
        - Each query result is written to a separate sheet in the Excel file. If a sheet name is not provided, 
          the sheets will be named as "Sheet1", "Sheet2", etc.
        - Logs are generated for connection status, query execution, and file writing.
    """
    config = read_config()
    user = config['database']['user']
    encrypted_password = config['database']['password']
    password = decrypt_password(encrypted_password)
    dsn = config['database']['dsn']

    queries = read_sql_queries(sql_file)

    try:
        # 使用 Thin 模式连接 Oracle 数据库
        connection = oracledb.connect(user=user, password=password, dsn=dsn)
        logging.info(f"Connected to Oracle database with DSN: {dsn}.")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for i, (sheet_name, query) in enumerate(queries):
                logging.info(f"Executing query: {query}")
                df = pd.read_sql(query, con=connection)
                sheet_name = sheet_name if sheet_name else f"Sheet{i+1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                logging.info(f"Query result written to sheet: {sheet_name}")
        logging.info(f"Data exported successfully to {output_file}")
        print(f"Data exported successfully to {output_file}")
    except Exception as e:
        logging.error(f"Failed to export data: {e}")
        print(f"Failed to export data: {e}")
        sys.exit(1)
    finally:
        if 'connection' in locals():
            connection.close()
            logging.info("Database connection closed.")

# 主函数
def main():
    if len(sys.argv) != 3:
        print("Usage: python oracle_report_batch.py <sql_file> <output_file>")
        logging.error("Invalid arguments. Usage: python oracle_report_batch.py <sql_file> <output_file>")
        sys.exit(1)

    sql_file = sys.argv[1]
    output_file = sys.argv[2]

    if not os.path.exists(sql_file):
        print(f"SQL file does not exist: {sql_file}")
        logging.error(f"SQL file does not exist: {sql_file}")
        sys.exit(1)

    export_to_excel(sql_file, output_file)

if __name__ == "__main__":
    main()