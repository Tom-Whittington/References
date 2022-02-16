import pyodbc
import matplotlib.pyplot as plt
import seaborn as sns
import sqlalchemy as sa
import pandas as pd

engine = sa.create_engine('mssql://GBWADJ8HJ293/ICP-OES?driver=SQL+Server+Native+Client+10.0?trusted_connection=yes')

df=pd.read_sql("SELECT * FROM TM21", engine)
print(df)
# df.to_csv("07Jul2021_TM21_Backup")





