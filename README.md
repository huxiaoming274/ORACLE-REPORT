# Oracle 数据导出工具

从 Oracle 数据库执行 SQL 查询，将结果导出为多 Sheet 的 Excel 文件。支持图形界面和命令行两种使用方式，密码加密存储。

---

## 功能特点

- 通过 SQL 文件定义查询，每条 SQL 对应 Excel 中的一个 Sheet
- 支持图形界面（GUI）和批量命令行两种模式
- 数据库密码使用 Fernet 对称加密，不明文存储
- 支持 **无需安装 Oracle Instant Client** 的 Thin 模式连接（`python-oracledb`）
- 也支持传统 `cx_Oracle` 模式（需要 Oracle Instant Client）

---

## 文件说明

### 主程序

| 文件 | 说明 |
|------|------|
| `export_data_encrypt_gui_v2.py` | **推荐** GUI 版，使用 `python-oracledb` Thin 模式，无需 Oracle Client，密码加密 |
| `export_data_encrypt_gui.py` | GUI 版，使用 `cx_Oracle`，密码加密 |
| `export_data_encrypt_gui_32bit.py` | GUI 版 32 位，使用 `cx_Oracle`，密码加密 |
| `export_data_gui.py` | GUI 版，使用 `cx_Oracle`，密码明文（早期版本） |
| `export_data_batch_without_oracle_client2.py` | **推荐** 命令行批量版，`python-oracledb` Thin 模式 |
| `export_data_batch_without_oracle_client.py` | 命令行批量版，`python-oracledb` Thin 模式 |
| `export_data_batch.py` | 命令行批量版，`cx_Oracle` |
| `encrypt_tool_gui.py` | 密码加密工具，用于生成 `secret.key` 并加密数据库密码 |

### 配置与 SQL

| 文件 | 说明 |
|------|------|
| `config.ini.example` | 配置文件模板，复制为 `config.ini` 并填写数据库信息 |
| `config.ini` | 实际配置文件（**不提交到 git**，包含凭证） |
| `secret.key` | Fernet 加密密钥（**不提交到 git**，由 `encrypt_tool_gui.py` 生成） |
| `11.sql` / `test.sql` | SQL 查询示例文件 |
| `getLastDayOfLastMonth.bat` | 获取上月最后一天日期的 Windows 批处理脚本 |

### 构建

| 文件 | 说明 |
|------|------|
| `*.spec` | PyInstaller 打包配置，用于生成独立可执行文件（`.exe`） |

---

## 快速开始

### 1. 安装依赖

```bash
pip install oracledb pandas openpyxl cryptography
```

> 若使用 `cx_Oracle` 版本，还需安装 Oracle Instant Client 并配置环境变量。

### 2. 配置数据库连接

```bash
cp config.ini.example config.ini
```

先用加密工具生成密钥并加密密码：

```bash
python encrypt_tool_gui.py
```

1. 点击"生成新的加密密钥"，生成 `secret.key`
2. 在输入框中填入数据库明文密码，点击"加密"
3. 将加密后的字符串填入 `config.ini` 的 `password` 字段

`config.ini` 格式：

```ini
[database]
user = YOUR_DB_USERNAME
password = YOUR_ENCRYPTED_PASSWORD
dsn = HOST:PORT/SERVICE_NAME

[output]
output_file = output.xlsx
```

> DSN 支持两种格式：`host:port/service_name` 或 `host:port:SID`

### 3. 编写 SQL 文件

每条 SQL 用 `;` 分隔，用注释指定对应的 Sheet 名：

```sql
-- sheet_name: 账户数据
select * from ACCOUNT1 t;

-- sheet_name: 债券数据
select * from CD_EMTN t;
```

### 4. 运行

**GUI 模式（推荐）：**

```bash
python export_data_encrypt_gui_v2.py
```

**命令行模式：**

```bash
python export_data_batch_without_oracle_client2.py query.sql output.xlsx
```

---

## 打包为可执行文件

使用 PyInstaller 和对应的 `.spec` 文件打包：

```bash
pyinstaller export_data_encrypt_gui_v2.spec
```

生成的 `.exe` 在 `dist/` 目录下，可在无 Python 环境的 Windows 机器上运行（需将 `config.ini` 和 `secret.key` 放在同目录）。

---

## 注意事项

- `config.ini` 和 `secret.key` 已加入 `.gitignore`，**不会提交到仓库**，请妥善保管
- `secret.key` 丢失后无法解密已加密的密码，需重新加密
- Excel Sheet 名称最长 31 个字符，超出会自动截断
