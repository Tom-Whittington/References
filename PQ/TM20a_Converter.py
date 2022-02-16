import pandas as pd
import os
import pyodbc

# Hg LOQ value
Hg_LOQ = 0.43

# DB connection string
conn = pyodbc.connect('Driver={SQL Server Native Client 10.0};'
                      'Server=GBWADJ8HJ293;'
                      'Database=ICP-OES;'
                      'Trusted_Connection=yes;')

# Data is imported with headers and footers removed.
df = pd.read_csv("importtemp.csv")

if os.path.exists("importtemp.csv"):
    os.remove("importtemp.csv")

# changes time column to datetime format
df["Time"] = pd.to_datetime(df["Time"])

# sets index to time
df.set_index = "Time"

# pivots calibration data and uses intensity as values
df_cal = pd.pivot(df, index=["Time", "Solution Label"], columns="Element", values="Int")

# slices calibration data
df_cal = df_cal[0:4]

# prints calibration data to csv
#df_cal_table.to_csv('TM20a_Cal.csv', mode='a', header=True)

# conc results are pivoted and indexed on time and solution label
df = pd.pivot(df, index=["Time", "Solution Label"], columns="Element", values="Corr Con")

# Slices sample data
df = df[5:-2]

# Blank correction
df = df - df.iloc[0]

# Drops blank
df = df[1:]

# Slices for control data
df_c = df.iloc[0]

# Transposes control
df_c = pd.DataFrame(df_c).T

# Prints control data to CSV
# df_c.to_csv("TM20a_Control.csv")

# Drops control
df = df.iloc[1:]

# Cu results for high copper products are set to zero
#df.loc[(df["Product"]=="BFG51") | (df["Product"] == "BFG52"), "Cu 327.395"] = 0

# THM calculated from sum of other elements
df["THM"] = df.sum(1)
print(df)
# Results are saved to CSV
# df.to_csv('TM20a_Trend.csv', mode='a', header=False)

# Other elements are dropped apart from THM and Hg
df = df[["THM", "Hg 184.887"]]

# Hg results less than than LOQ (0.43) are changed to <LOQ value
df.loc[df["Hg 184.887"] < Hg_LOQ, "Hg 184.887"] = "<0.43"

# Index is reset
df.reset_index(inplace=True)

# Time column is dropped after trending is complete
df = df.drop("Time", axis=1)

# df is melted and all values are put under result.resultentry column ready for LIMS import
df = pd.melt(df, id_vars="Solution Label", var_name="Element", value_name="result.resultentry")
print(df)

# Spec information is requested from DB
SQL = "SELECT * FROM TM20a_Spec"
df_ref = pd.read_sql(SQL, conn)

# df_ref["ProdEle] is new column created by merging "Product and "Element" columns
df_ref["ProdEle"] = df_ref["Product"] + " " + df_ref["Element"]

# Old columns are dropped
df_ref.drop(["Product", "Element"], axis=1, inplace=True)
print(df_ref)
