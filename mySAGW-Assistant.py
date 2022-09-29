# import
# !/usr/local/bin/python3
import sys
from tkinter.filedialog import askdirectory
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# initialisieren
w = 0
x = 0
y = 0
z = 0
n_hm = True

# Willkommensmeldung
print()
print("Guten Tag,")
print("Willkommen beim digitalen mySAGW-Assistenten!")
print("Bitte wählen Sie die Dateien aus (Mehrfachauswahl möglich)")
print()

# Hauptprogramm
# 1 Separate Listen, aufgeteilt nach Sektionen
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilenames(title="Dateien öffnen (Mehrfachauswahl möglich)")
print(file_path)
a = len(file_path)
print(a)

print()
print("Wo soll der Output gespeichert werden?")
path = askdirectory(title='Speicherort auswählen')
print()

df = pd.DataFrame(pd.read_excel(file_path[0]))
print(df)
print()

for element in file_path[1:9]:
    dftemp = pd.DataFrame(pd.read_excel(element))
    frames = [df, dftemp]
    df = pd.concat(frames)
    print(df)
    print()

print(df)
print()

# Sektion 1
b = 1
sektion1 = "sektion_" + str(b) + ".xlsx"
dfsek1 = df[(df.SEKTION == "Sektion 1: Historische und archäologische Wissenschaften")]
dfsek1 = dfsek1.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
dfsek1.to_excel(path + "/" + sektion1)

# Sektion 2
b += 1
sektion2 = "sektion_" + str(b) + ".xlsx"
dfsek2 = df[(df.SEKTION == "Sektion 2: Kunstwissenschaften")]
dfsek2 = dfsek2.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
dfsek2.to_excel(path + "/" + sektion2)

# Sektion 3
b += 1
sektion3 = "sektion_" + str(b) + ".xlsx"
dfsek3 = df[(df.SEKTION == "Sektion 3: Sprach- und Literaturwissenschaften")]
dfsek3 = dfsek3.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
dfsek3.to_excel(path + "/" + sektion3)

# Sektion 4
b += 1
sektion4 = "sektion_" + str(b) + ".xlsx"
dfsek4 = df[(df.SEKTION == "Sektion 4: Kulturwissenschaften")]
dfsek4 = dfsek4.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
dfsek4.to_excel(path + "/" + sektion4)

# Sektion 5
b += 1
sektion5 = "sektion_" + str(b) + ".xlsx"
dfsek5 = df[(df.SEKTION == "Sektion 5: Wirtschafts- und Rechtswissenschaften")]
dfsek5 = dfsek5.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
dfsek5.to_excel(path + "/" + sektion5)

# Sektion 6
b += 1
sektion6 = "sektion_" + str(b) + ".xlsx"
dfsek6 = df[(df.SEKTION == "Sektion 6: Gesellschaftswissenschaften")]
dfsek6 = dfsek6.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
dfsek6.to_excel(path + "/" + sektion6)

# Sektion 7
b += 1
sektion7 = "sektion_" + str(b) + ".xlsx"
dfsek7 = df[(df.SEKTION == "Sektion 7: Wissenschaft – Technik – Gesellschaft")]
dfsek7 = dfsek7.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
dfsek7.to_excel(path + "/" + sektion7)

# Kommissionen / Kuratorien
komm = "kommissionen_kuratorien.xlsx"
dfsek7 = df[(df.SEKTION == "Kommission / Kuratorium")]
dfsek7 = dfsek7.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
dfsek7.to_excel(path + "/" + komm)

print()
print("Speichern abgeschlossen!")
print()
print("Herzlichen Dank für Ihren Besuch und bis bald!")
sys.exit(0)
