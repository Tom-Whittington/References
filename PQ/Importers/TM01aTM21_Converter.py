import pandas as pd
import numpy as np
import pyodbc
import os

# Stores LOQ values
loq = 0, 0.25, 0, 0.06, 0.19, 0.25, 0.09, 0, 0, 0.27, 0.17, 0.18, 0, 0.11, 0

# Converted LOQ values into np array
loq_arr = np.array([loq])

# Converts each LOQ value to string and adds "<"
loq_str = ["<" + str(loq) for loq in loq]

# Creates np array of "<" LOQ values
loq_str_arr = np.array([loq_str])

# DB connection string
conn = pyodbc.connect('Driver={SQL Server Native Client 10.0};'
                      'Server=GBWADJ8HJ293;'
                      'Database=ICP-OES;'
                      'Trusted_Connection=yes;')


# todo convert extra columns to sample number, product etc then create raw data database, then upload raw data to
#  database


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
        elif np.greater(x,0.01):
            return round(x, 2)
        # else round to 3 dp
        else:
            return round(x, 3)
        # if element is not float then return element
    else:
        return x


# imports data
df = pd.read_csv("importtemp.csv")

# deletes import file
if os.path.exists("importtemp.csv"):
    os.remove("importtemp.csv")

# changes time column to datetime format
df["Time"] = pd.to_datetime(df["Time"])

# sets index to time
df.set_index = "Time"

# slices for TM21 samples based on weight being greater than 4.5 g. Excludes EUR samples (usually 6 g)
filt = (df["Act Wgt"] > 4.5) & (~df["Solution Label"].str.contains("EUR", na=False))
df_TM21 = df[filt]

# Checks if there are any TM21 samples, if not it skips over
if not df_TM21.empty:
    # pivots TM21 df setting index on time and solution label
    df_TM21 = pd.pivot(df_TM21, index=["Time", "Solution Label"], columns="Element", values="Corr Con")

    # TM21 blank correction
    df_TM21 = df_TM21 - df_TM21.iloc[0]

    # drops blank
    df_TM21 = df_TM21.iloc[1:]

    # sets control
    df_TM21_c = df_TM21.iloc[0]

    # transposes control value
    df_TM21_c = pd.DataFrame(df_TM21_c).T

    # prints control data to csv
    # df_TM21_c.to_csv("TM21_Control.csv")

    # drops control
    df_TM21 = df_TM21.iloc[1:]

    # prints trend data to csv
    # df_TM21.to_csv("TM21_Trend.csv")

    # counts how many TM21 samples there are
    rows = len(df_TM21.index)

    # generates LOQ array of same size as sample size
    loq_str_arr = np.repeat(a=loq_str_arr, repeats=rows, axis=0)

    # If values aren't above LOQ then they are replaced with <LOQ
    df_TM21[df_TM21 < loq_arr] = loq_str_arr

    # resets index
    df_TM21.reset_index(inplace=True)

    # drops time after trending is complete
    df_TM21 = df_TM21.drop("Time", axis=1)

    # df is melted and all values are put under result.resultentrycolumn ready for LIMS import
    df_TM21 = pd.melt(df_TM21, id_vars="Solution Label", var_name="Element", value_name="result.resultentry")

    # results are formatted
    df_TM21["result.resultentry"] = df_TM21["result.resultentry"].apply(lambda x: formatter(x))

    # todo prepare df_TM21 for merge (need sample no. and product)

    # TM21 Spec information is requested from DB
    SQL = "SELECT * FROM TM21_Spec"
    df_ref = pd.read_sql(SQL, conn)

    # df_ref["ProdEle] is a new column created by merging "Product and "Element" columns
    df_ref["ProdEle"] = df_ref["Product"] + " " + df_ref["Element"]

    # Old columns are dropped
    df_ref.drop(["Product", "Element"], axis=1, inplace=True)

    # df and df_ref are merged to add LIMS tests to df. Inner join, any products/test combinations not found in
    # df_ref are dropped from df
    df_TM21 = pd.merge(df_TM21, df_ref)

# slices out TM21 results

filt = (df["Act Wgt"] < 4.5) | (df["Solution Label"].str.contains("EUR", na=False))
df = df[filt]

# pivots calibration data and uses intensity as values
df_cal = pd.pivot(df, index=["Time", "Solution Label"], columns="Element", values="Int")

# slices calibration data
df_cal = df_cal[0:6]

# prints calibration data to csv
# df_cal_table.to_csv('TM01a_Cal.csv', mode='a', header=False)

# conc results are pivoted and indexed on time and solution label
df = pd.pivot(df, index=["Time", "Solution Label"], columns="Element", values="Corr Con")

# Slices sample data
df = df[7:-4]

# Blank correction
df = df - df.iloc[0]

# Drops blank
df = df[1:]

# Slices for control data
df_c = df.iloc[0]

# Transposes control
df_c = pd.DataFrame(df_c).T

# Prints control data to CSV
# df_c.to_csv("TM01a_Control.csv")

# Drops control
df = df.iloc[1:]

# Al converted to Al2O3
df["Al 396.152"] = df["Al 396.152"] * 1.89

# Prints trending data to CSV
# df.to_csv('TM01a_Trend.csv', mode='a', header=True)

# Counts how many samples in df
rows = len(df.index)

# Creates np array of LOQ values. Values are repeated the same amount of time as there are samples
loq_str_arr = np.repeat(a=loq_str_arr, repeats=rows, axis=0)

# If values aren't above LOQ then they are replaced with <LOQ
df[df < loq_arr] = loq_str_arr

# Index is reset
df.reset_index(inplace=True)

# Time column is dropped after trending is complete
df = df.drop("Time", axis=1)

# df is melted and all values are put under result.resultentry column ready for LIMS import
df = pd.melt(df, id_vars="Solution Label", var_name="Element", value_name="result.resultentry")

# applies number formatting
df["result.resultentry"] = df["result.resultentry"].apply(lambda x: formatter(x))

# todo prepare df_TM01a for merge (need sample no. and product)

# Spec information is requested from DB
SQL = "SELECT * FROM TM01a_Spec"
df_ref = pd.read_sql(SQL, conn)

# df_ref["ProdEle] is new column created by merging "Product and "Element" columns
df_ref["ProdEle"] = df_ref["Product"] + " " + df_ref["Element"]

# Old columns are dropped
df_ref.drop(["Product", "Element"], axis=1, inplace=True)

# df and df_ref are merged to add LIMS tests to df. Inner join, any products/test combinations not found in df_ref are
# dropped from df
# df = pd.merge(df, df_ref)

# todo add division by 10000 if "cent" is in LIMS test
