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
        path_sek = path + "/" + sektion
        if os.path.exists(path_sek):
            os.remove(path_sek)
            print("File " + sektion + " deleted.")
            dfsek.to_excel(path + "/" + sektion)
            print("File "  + sektion + " created.")
            print()
            b += 1
        else:
            print("File "  + sektion + " doesn't exist in this location.")
            dfsek.to_excel(path + "/" + sektion)
            print("File "  + sektion + " created.")
            print()
            b += 1
    else:
        komm = "kommissionen_kuratorien.xlsx"
        dfkomm = df[(df.SEKTION == "Kommission / Kuratorium")]
        dfkomm = dfkomm.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
        path_komm = path + "/" + komm
        if os.path.exists(path_komm):
            os.remove(path_komm)
            print("File " + komm + " deleted.")
            dfkomm.to_excel(path + "/" + komm)
            print("File "  + komm + " created.")
            print()
        else:
            print("File "  + komm + " doesn't exist in this location.")
            dfkomm.to_excel(path + "/" + komm)
            print("File "  + komm + " created.")
            print()

print()
print("Files saved!")
print()
print("Thank you for using the mySAGW-Assistant-Application!")
sys.exit(0)