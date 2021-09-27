"""
MySQLへのデータ書き込み
No module named 'MySQLdb'が出たら
conda install mysqlclient 失敗
1. pip install mysqlclient
2. Mysqlサーバーでnew_dbのデータベースを作成　パスワードも更新
2021/3/24 (水) 16:03
"""
from sqlalchemy import create_engine
import pandas as pd
import pandas.io.sql as psql

table_name = "jpstock"
db_settings = {
    "host": 'localhost',
    "database": 'new_db',
    "user": 'root',
    "password": '2021',
    "port":3306
}
engine = create_engine('mysql://{user}:{password}@{host}:{port}/{database}?charset=utf8'.format(**db_settings))


df = pd.DataFrame([['sample3', 'CCC', '2017-07-16 00:00:00']], columns=['title', 'body', 'created'], index=[2])

with engine.begin() as con:
    df.to_sql(table_name, con=con, if_exists='append', index=False)
