import pandas as pd
import os
import pyodbc
import numpy as np


def formatter(x):
    """Applies number formatting to data frame values based on value"""
    # if df element is number
    if type(x) == float:
        # if element is greater than 100, round to 0 dp
        if np.greater(x, 100):
            return round(x, 0)
        # if element is greater than 10, round to 1 dp
        elif np.greater(x, 10):
            return round(x, 1)
        # if element is greater than 0.01, round to 2 dp
        elif np.greater(x, 0.01):
            return round(x, 2)
        # else round to 3 dp
        else:
            return round(x, 3)
        # if element is not float then return element
    else:
        return x


# DB connection string
conn = pyodbc.connect('Driver={SQL Server Native Client 10.0};'
                      'Server=GBWADJ8HJ293;'
                      'Database=ICP-OES;'
                      'Trusted_Connection=yes;')

# data import
df = pd.read_csv("importtemp.csv")

# deletes import file
if os.path.exists("importtemp.csv"):
    os.remove("importtemp.csv")

# changes time column to datetime format
df["Time"] = pd.to_datetime(df["Time"])

# sets index to time
df.set_index = "Time"

# pivots calibration data and uses intensity as values
df_cal = pd.pivot(df, index=["Time", "Solution Label"], columns="Element", values="Int")

# slices calibration data
df_cal = df_cal[0:5]

# prints calibration data to csv
# df_cal.to_csv('TM01g_Cal.csv', mode='a', header=True)

# conc results are pivoted and indexed on time and solution label
df = pd.pivot(df, index=["Time", "Solution Label"], columns="Element", values="Corr Con")

# Slices sample data
df = df[6:-1]

# Blank correction
df = df - df.iloc[0]

# Drops blank
df = df[1:]

# Slices for control data
df_c = df.iloc[0]

# Transposes control
df_c = pd.DataFrame(df_c).T

# Prints control data to CSV
# df_c.to_csv("TM01g_Control.csv")

# Drops control
df = df.iloc[1:]

# Al converted to Al2O3
df["Al 396.152"] = df["Al 396.152"] * 1.89

# Trend results sent to CSV
# df.to_csv('TM01g_Trend.csv', mode='a', header=False, index=True)

# Index is reset
df.reset_index(inplace=True)

# Time column is dropped after trending is complete
df = df.drop("Time", axis=1)

# df is melted and all values are put under result.resultentry column ready for LIMS import
df = pd.melt(df, id_vars="Solution Label", var_name="Element", value_name="result.resultentry")


# applies number formatting
df["result.resultentry"] = df["result.resultentry"].apply(lambda x: formatter(x))

# todo prepare df_TM01g for merge (need sample no. and product)

# Spec information is requested from DB
SQL = "SELECT * FROM TM01g_Spec"
df_ref = pd.read_sql(SQL, conn)

print(df_ref)
print(df)
# df_ref["ProdEle] is new column created by merging "Product and "Element" columns
df_ref["ProdEle"] = df_ref["Product"] + " " + df_ref["Element"]

# Old columns are dropped
df_ref.drop(["Product", "Element"], axis=1, inplace=True)

# df and df_ref are merged to add LIMS tests to df. Inner join, any products/test combinations not found in df_ref are
# dropped from df
# df = pd.merge(df, df_ref)
