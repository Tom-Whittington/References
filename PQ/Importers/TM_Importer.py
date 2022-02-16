import pandas as pd
import numpy as np
import pyodbc
from tkinter.filedialog import askopenfilename


def importer():
    """ imports csv data, calls variables to get run information and slices out TM21 data"""
    global df_imp, method, smp_end
    filename = askopenfilename()
    df_imp = pd.read_csv(filename, header=3)

    # Cuts out footer (footer changes position due to audit trail below results)
    i = df_imp.index.get_loc(df_imp.index[df_imp["Solution Label"] == "Worksheet Signature Trail"][0])
    df_imp = df_imp[:i]

    variables(df_imp)

    # Creates datetime column
    df_imp["DateTime"] = df_imp["Date"] + " " + df_imp["Time"]
    df_imp["DateTime"] = pd.to_datetime(df_imp["DateTime"])

    # sets index to datetime
    df_imp.set_index = "DateTime"
    df_cal = pd.pivot(df_imp, index=["DateTime", "Solution Label"], columns="Element", values="Int")

    # slices calibration data
    df_cal = df_cal[:cal_end]

    # prints calibration data to csv
    df_cal.to_csv(method + '_Cal.csv', mode='a', header=False)

    if method == "TM01a":
        # slices for TM21 samples based on weight being greater than 4.5 g. Excludes EUR samples (usually 6 g)
        filt = (df_imp["Act Wgt"] > 4.5) & (~df_imp["Solution Label"].str.contains("EUR", na=False))
        df = df_imp[filt]

        # checks for TM21 samples
        if not df.empty:
            method = "TM21"

            # calls processor on TM21 samples
            processor(df, loq_arr, loq_str_arr)

            # reverts back to TM01a method
            method = "TM01a"

            # accounts for the extra wash injection
            smp_end = smp_end - 1

            # slices for other injections
            filt = (df_imp["Act Wgt"] < 4.5) | (df_imp["Solution Label"].str.contains("EUR", na=False))
            df_imp = df_imp[filt]

        # calls processor on TM01a samples
        processor(df_imp, loq_arr, loq_str_arr)

    else:
        processor(df_imp, loq_arr, loq_str_arr)


def variables(df_imp):
    """counts number of unique elements in import file and sets method parameters"""
    global method, loq_arr, cal_end, smp_end, loq_str_arr

    # counts number of unique elements tested
    elements = df_imp["Element"].nunique()

    if elements == 15:
        method = "TM01a"
        loq = 0, 0.25, 0, 0.06, 0.19, 0.25, 0.09, 0, 0, 0.27, 0.17, 0.18, 0, 0.11, 0
        cal_end = 6
        smp_end = -3

    elif elements == 13:
        method = "TM20a"
        loq = 0, 0.43
        cal_end = 4
        smp_end = -2

    else:
        loq = 0, 0
        method = "TM01g"
        cal_end = 5
        smp_end = -2

    # Converted LOQ values into np array
    loq_arr = np.array([loq])

    # Converts each LOQ value to string and adds "<"
    loq_str = ["<" + str(loq) for loq in loq]

    # Creates np array of "<" LOQ values
    loq_str_arr = np.array([loq_str])


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


def processor(df_imp, loq_arr, loq_str_arr):
    """Processes raw data into usable, saves trending data and finally joins with LIMS spec for upload"""
    # imports data
    df = pd.pivot(df_imp, index=["DateTime", "Solution Label"], columns="Element", values="Corr Con")

    # checks for TM21
    if method != "TM21":

        # Slices sample data
        df = df[cal_end + 1:smp_end]

    # Blank correction
    df = df - df.iloc[0]

    # Drops blank
    df = df[1:]

    # Slices for control data
    df_c = df.iloc[0]

    # Transposes control
    df_c = pd.DataFrame(df_c).T

    # Prints control data to CSV
    df_c.to_csv(method + "_Control.csv")

    # Drops control
    df = df.iloc[1:]

    # If df isn't TM20a then Al needs to be corrected to Al2O3
    if method != "TM20a":
        df["Al 396.152"] = df["Al 396.152"] * 1.89

    else:
        # THM calculated from sum of other elements
        df["THM_TM20a"] = df.sum(1)

    # Results are saved to CSV
    df.to_csv(method + " Trend.csv", mode='a', header=True)

    # Other elements are dropped apart from THM and Hg
    if method == "TM20a":
        df = df[["THM", "Hg 184.887"]]

    # samples are counted
    rows = len(df.index)

    # Creates np array of LOQ values. Values are repeated the same amount of time as there are samples
    loq_str_arr = np.repeat(a=loq_str_arr, repeats=rows, axis=0)

    # If values aren't above LOQ then they are replaced with <LOQ
    if method != "TM01g":
        df[df < loq_arr] = loq_str_arr

    # Index is reset
    df.reset_index(inplace=True)

    # Time column is dropped after trending is complete
    df = df.drop("DateTime", axis=1)

    # df is melted and all values are put under result.resultentry column ready for LIMS import
    df = pd.melt(df, id_vars="Solution Label", var_name="Element", value_name="result.resultentry")
    df['Element'] = df['Element'].str.replace(r" \d+.\d+", "_" + method, regex=True)

    # applies number formatting
    df["result.resultentry"] = df["result.resultentry"].apply(lambda x: formatter(x))

    # DB connection string
    conn = pyodbc.connect('Driver={SQL Server Native Client 10.0};'
                          'Server=GBWADJ8HJ293;'
                          'Database=ICP-OES;'
                          'Trusted_Connection=yes;')

    # Spec information is requested from DB
    sql = "SELECT * FROM TM_Spec"
    df_ref = pd.read_sql(sql, conn)

    # df_ref["ProdEle] is new column created by merging "Product and "Element" columns
    df_ref["ProdEle"] = df_ref["Product"] + " " + df_ref["Element"]

    # Old columns are dropped
    df_ref.drop(["Product", "Element"], axis=1, inplace=True)
    print(df_ref)
    print(df)
    df.to_csv(method + " test.csv")

    cent = df['test_analysis'].str.contains("cent") == 'euro'

    # Converts ppm results into %
    df.loc[cent, "result.resultentry"] = df.loc[cent, 'result.resultentry'] / 10000


importer()
