import pandas as pd
from tkinter.filedialog import askopenfilename

filename = askopenfilename()
df = pd.read_csv(filename, header=3, skipfooter=4, engine='python',)

elements=df["Element"].nunique()

df.to_csv("importtemp.csv", index=False)

if elements == 15:
    import TM01aTM21_Converter

elif elements == 13:
    import TM20a_Converter

else:
    import TM01g_Converter