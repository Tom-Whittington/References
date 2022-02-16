import pyodbc
import matplotlib.pyplot as plt
import seaborn as sns
import sqlalchemy as sa
import pandas as pd
import os
#
# rootdir=r'W:\Analytical Group\Control Charts\ICP - 720\TM Sample Raw Data\TM21\2021'
# for subdir, dirs, files in os.walk(rootdir):
#     for file in files:
#         # print os.path.join(subdir, file)
#         filepath = subdir + os.sep + file
#
#         if filepath.endswith(".xlsx"):
#             # Data is imported with headers and footers removed.
#             df = pd.read_excel(filepath)
#
#             df.to_csv(r'C:\Users\tom.whittington\Documents\ExcelTM21.csv', mode='a', header=False, index=False)
#             continue
#
#
# conn = pyodbc.connect('Driver={SQL Server Native Client 10.0};'
#                       'Server=GBWADJ8HJ293;'
#                       'Database=ICP-OES;'
#                       'Trusted_Connection=yes;')
#
#
#
engine = sa.create_engine('mssql+pyodbc://GBWADJ8HJ293/ICP-OES?driver=SQL+Server+Native+Client+10.0?trusted_connection=yes')

# df=pd.read_excel(r"C:\Users\tom.whittington\Documents\ExcelTM21.xlsx")
# df["Date"] = pd.to_datetime(df["Date"]).dt.strftime('%Y-%m-%d')
#
# df.to_sql("TM21",engine, if_exists="append", index=False)
df=pd.read_sql("SELECT * FROM RAW DATA", engine)
print(df)




# engine.execute("DELETE FROM TM20a WHERE Solution_Label= '' ")



