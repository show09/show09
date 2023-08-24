import sqlite3
import pandas as pd

# 连接到SQLite数据库
db_connection = sqlite3.connect('your_database.db')  # 替换为你的数据库文件路径
cursor = db_connection.cursor()

# 从SQL文件中读取SQL语句
with open('queries.sql', 'r') as sql_file:
    sql_queries = sql_file.read()

# 执行SQL查询
cursor.executescript(sql_queries)
db_connection.commit()

# 从数据库中检索数据并存储在DataFrame中
query_result = cursor.execute('SELECT * FROM your_table')  # 替换为你的表名和查询
data = query_result.fetchall()
column_names = [description[0] for description in query_result.description]
df = pd.DataFrame(data, columns=column_names)

# 将数据导出到Excel文件
df.to_excel('output.xlsx', index=False, engine='openpyxl')  # 替换为你想要的输出文件名

# 关闭数据库连接
db_connection.close()
