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
z = "*"

# Willkommensmeldung
print()
print("Guten Tag,")
print("Willkommen beim digitalen mySAGW-Assistenten!")
print("Bitte wählen Sie die Dateien aus (Mehrfachauswahl möglich)")
print()
print(25*z)
print("Achtung: Bitte sicherstellen, dass die nachfolgenden Spaltenüberschriften in Grossbuchstaben vorliegen:")
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
file_path = filedialog.askopenfilenames(title="Dateien öffnen (Mehrfachauswahl möglich)")
a = len(file_path)
if file_path == "":
    print("Keine Datei ausgewählt, Programm wird beendet")
    sys.exit(0)

print()
print("Wo soll der Output gespeichert werden?")
path = askdirectory(title='Speicherort auswählen')
if path == "":
    print("Kein Speicherort ausgewählt, Programm wird beendet")

if a > 1:
    df = pd.DataFrame(pd.read_excel(file_path[0]))
    for element in file_path[1:9]:
        dftemp = pd.DataFrame(pd.read_excel(element))
        frames = [df, dftemp]
        df = pd.concat(frames)

else:
    for element in file_path:
        df = pd.DataFrame(pd.read_excel(element))

# Sektion 1
b = 1
found = df[df['SEKTION'].str.contains('Sektion 1: Historische und archäologische Wissenschaften')]
v = len(found)

if v > 0:
    sektion1 = "sektion_" + str(b) + ".xlsx"
    dfsek1 = df[(df.SEKTION == "Sektion 1: Historische und archäologische Wissenschaften")]
    dfsek1 = dfsek1.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
    dfsek1.to_excel(path + "/" + sektion1)
else:
    pass

# Sektion 2
b += 1
try:
    sektion2 = "sektion_" + str(b) + ".xlsx"
    dfsek2 = df[(df.SEKTION == "Sektion 2: Kunstwissenschaften")]
    dfsek2 = dfsek2.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
    dfsek2.to_excel(path + "/" + sektion2)
except:
    pass

# Sektion 3
b += 1
try:
    sektion3 = "sektion_" + str(b) + ".xlsx"
    dfsek3 = df[(df.SEKTION == "Sektion 3: Sprach- und Literaturwissenschaften")]
    dfsek3 = dfsek3.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
    dfsek3.to_excel(path + "/" + sektion3)
except:
    pass

# Sektion 4
b += 1
try:
    sektion4 = "sektion_" + str(b) + ".xlsx"
    dfsek4 = df[(df.SEKTION == "Sektion 4: Kulturwissenschaften")]
    dfsek4 = dfsek4.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
    dfsek4.to_excel(path + "/" + sektion4)
except:
    pass

# Sektion 5
b += 1
try:
    sektion5 = "sektion_" + str(b) + ".xlsx"
    dfsek5 = df[(df.SEKTION == "Sektion 5: Wirtschafts- und Rechtswissenschaften")]
    dfsek5 = dfsek5.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
    dfsek5.to_excel(path + "/" + sektion5)
except:
    pass

# Sektion 6
b += 1
try:
    sektion6 = "sektion_" + str(b) + ".xlsx"
    dfsek6 = df[(df.SEKTION == "Sektion 6: Gesellschaftswissenschaften")]
    dfsek6 = dfsek6.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
    dfsek6.to_excel(path + "/" + sektion6)
except:
    pass

# Sektion 7
b += 1
try:
    sektion7 = "sektion_" + str(b) + ".xlsx"
    dfsek7 = df[(df.SEKTION == "Sektion 7: Wissenschaft – Technik – Gesellschaft")]
    dfsek7 = dfsek7.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
    dfsek7.to_excel(path + "/" + sektion7)
except:
    pass

# Kommissionen / Kuratorien
try:
    komm = "kommissionen_kuratorien.xlsx"
    dfsek7 = df[(df.SEKTION == "Kommission / Kuratorium")]
    dfsek7 = dfsek7.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
    dfsek7.to_excel(path + "/" + komm)
except:
    pass

print()
print("Speichern abgeschlossen!")
print()
print("Herzlichen Dank für Ihren Besuch und bis bald!")
sys.exit(0)
