import pandas as pd
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename

# Read in the Excel file
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file

df = pd.read_excel(filename)

# replace NaN values with 0
df.fillna(value=0, inplace=True)

# initialize a new dataframe to store the results
result = pd.DataFrame(columns=df.columns)

# iterate over the rows of the original dataframe
subtotal = pd.Series([0]*(len(df.columns)), index=df.columns)
last_value = None
for i, row in df.iterrows():
    if last_value != row["MITGLIEDINSTITUTION"]:
        # new value in Column A, add subtotal row
        if i != 0:
            subtotal["MITGLIEDINSTITUTION"] = "Subtotal"
            subtotal["VERTEILPLAN"] = ""
            subtotal["SEKTION"] = ""
            subtotal["FORM_ID"] = ""
            subtotal["REFERENZ-NR."] = ""
            subtotal["PROJEKTTITEL"] = ""
            subtotal["GESUCHSTYP"] = ""
            result = result.append(subtotal, ignore_index=True)
        subtotal = pd.Series([row["MITGLIEDINSTITUTION"]] + [0]*(len(df.columns)-1), index=df.columns)
        last_value = row["MITGLIEDINSTITUTION"]

    # add current row to result and subtotal
    result = result.append(row)
    subtotal["BEANTRAGT"] += row["BEANTRAGT"]
    subtotal["VORSCHLAG WIMA"] += row["VORSCHLAG WIMA"]
    subtotal["GESPROCHEN"] += row["GESPROCHEN"]
    subtotal["VORSCHUSSZAHLUNG"] += row["VORSCHUSSZAHLUNG"]
    subtotal["AUSZAHLUNGSBETRAG"] += row["AUSZAHLUNGSBETRAG"]
    subtotal["VERTEILPLAN"] = ""
    subtotal["SEKTION"] = ""
    subtotal["FORM_ID"] = ""
    subtotal["REFERENZ-NR."] = ""
    subtotal["PROJEKTTITEL"] = ""
    subtotal["GESUCHSTYP"] = ""

# add final subtotal row
subtotal["MITGLIEDINSTITUTION"] = "Subtotal"
subtotal["VERTEILPLAN"] = ""
subtotal["SEKTION"] = ""
subtotal["FORM_ID"] = ""
subtotal["REFERENZ-NR."] = ""
subtotal["PROJEKTTITEL"] = ""
subtotal["GESUCHSTYP"] = ""
result = result.append(subtotal, ignore_index=True)

# write the results to a new Excel file
result.to_excel("output_file.xlsx", index=False)
