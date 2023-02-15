# import
# !/usr/local/bin/python3
import sys
from tkinter.filedialog import askdirectory
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# initialisieren
b = 1
w = 0
x = 0
y = 0
z = "*"

sektion_list = ['Sektion 1: Historische und archäologische Wissenschaften',
                'Sektion 2: Kunstwissenschaften',
                'Sektion 3: Sprach- und Literaturwissenschaften',
                'Sektion 4: Kulturwissenschaften',
                'Sektion 5: Wirtschafts- und Rechtswissenschaften',
                'Sektion 6: Gesellschaftswissenschaften',
                'Sektion 7: Wissenschaft – Technik – Gesellschaft',
                'Kommission / Kuratorium']

# Willkommensmeldung
print()
print("Hello,")
print("Welcome to the mySAGW-Assistant!")
print("Please choose the corresponding files (multi-selection possible)")
print()
print(25*z)
print("Warning: Please ensure to use the correct spelling of the following columns:")
print()
print("SEKTION")
print("MITGLIEDINSTITUTION")
print("FORM_ID")
print("REFERENZ-NR.")
print(25*z)

# Hauptprogramm
# 1 Separate Listen, aufgeteilt nach Sektionen
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilenames(title="Open File(s) (Multi-Selection Possible)")
a = len(file_path)
if file_path == "":
    print()
    print("No file selected, closing application...")
    print()
    sys.exit(0)

print()
print("Where do you wish to save the output?")
path = askdirectory(title='Choose save location')
if path == "":
    print("No save location selected, closing application...")
print()

if a > 1:
    df = pd.DataFrame(pd.read_excel(file_path[0]))
    for element in file_path[1:9]:
        dftemp = pd.DataFrame(pd.read_excel(element))
        frames = [df, dftemp]
        df = pd.concat(frames)

else:
    for element in file_path:
        df = pd.DataFrame(pd.read_excel(element))

for elem in sektion_list:
    if elem != "Kommission / Kuratorium":
        sektion = "sektion_" + str(b) + ".xlsx"
        dfsek = df[(df.SEKTION == elem)]
        dfsek = dfsek.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])

        # replace NaN values with 0
        dfsek.fillna(value=0, inplace=True)

        # initialize a new dataframe to store the results
        dfSekResult = pd.DataFrame(columns=dfsek.columns)

        # iterate over the rows of the original dataframe
        subtotal = pd.Series([0]*(len(dfsek.columns)), index=dfsek.columns)
        last_value = None
        for i, row in dfsek.iterrows():
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
                    dfSekResult = dfSekResult.append(subtotal, ignore_index=True)
                subtotal = pd.Series([row["MITGLIEDINSTITUTION"]] + [0]*(len(dfsek.columns)-1), index=dfsek.columns)
                last_value = row["MITGLIEDINSTITUTION"]

            # add current row to result and subtotal
            dfSekResult = dfSekResult.append(row)
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
        dfSekResult = dfSekResult.append(subtotal, ignore_index=True)

        path_sek = path + "/" + sektion
        if os.path.exists(path_sek):
            os.remove(path_sek)
            print("File " + sektion + " deleted.")
            dfSekResult.to_excel(path + "/" + sektion, index=False)
            print("File "  + sektion + " created.")
            print()
            b += 1
        else:
            print("File "  + sektion + " doesn't exist in this location.")
            dfSekResult.to_excel(path + "/" + sektion, index=False)
            print("File "  + sektion + " created.")
            print()
            b += 1
    else:
        komm = "kommissionen_kuratorien.xlsx"
        dfkomm = df[(df.SEKTION == "Kommission / Kuratorium")]
        dfkomm = dfkomm.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])

        # replace NaN values with 0
        dfkomm.fillna(value=0, inplace=True)

        # initialize a new dataframe to store the results
        dfKommResult = pd.DataFrame(columns=dfkomm.columns)

        # iterate over the rows of the original dataframe
        subtotal = pd.Series([0]*(len(dfkomm.columns)), index=dfkomm.columns)
        last_value = None
        for i, row in dfkomm.iterrows():
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
                    dfKommResult = dfKommResult.append(subtotal, ignore_index=True)
                subtotal = pd.Series([row["MITGLIEDINSTITUTION"]] + [0]*(len(dfkomm.columns)-1), index=dfkomm.columns)
                last_value = row["MITGLIEDINSTITUTION"]

            # add current row to result and subtotal
            dfKommResult = dfKommResult.append(row)
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
        dfKommResult = dfKommResult.append(subtotal, ignore_index=True)

        path_komm = path + "/" + komm
        if os.path.exists(path_komm):
            os.remove(path_komm)
            print("File " + komm + " deleted.")
            dfKommResult.to_excel(path + "/" + komm, index=False)
            print("File "  + komm + " created.")
            print()
        else:
            print("File "  + komm + " doesn't exist in this location.")
            dfKommResult.to_excel(path + "/" + komm, index=False)
            print("File "  + komm + " created.")
            print()

print()
print("Files saved!")
print()
print("Thank you for using the mySAGW-Assistant-Application!")
sys.exit(0)