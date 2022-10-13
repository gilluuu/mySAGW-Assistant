# import
# !/usr/local/bin/python3
import sys
from tkinter.filedialog import askdirectory
import os
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

# Sektion 1
b = 1
found = df[df['SEKTION'].str.contains('Sektion 1: Historische und archäologische Wissenschaften')]
v = len(found)

sektion1 = "sektion_" + str(b) + ".xlsx"
dfsek1 = df[(df.SEKTION == "Sektion 1: Historische und archäologische Wissenschaften")]
dfsek1 = dfsek1.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
path_sek1 = path + "/" + sektion1


# Sektion 2
b += 1
sektion2 = "sektion_" + str(b) + ".xlsx"
dfsek2 = df[(df.SEKTION == "Sektion 2: Kunstwissenschaften")]
dfsek2 = dfsek2.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
path_sek2 = path + "/" + sektion2


# Sektion 3
b += 1
sektion3 = "sektion_" + str(b) + ".xlsx"
dfsek3 = df[(df.SEKTION == "Sektion 3: Sprach- und Literaturwissenschaften")]
dfsek3 = dfsek3.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
path_sek3 = path + "/" + sektion3

# Sektion 4
b += 1
sektion4 = "sektion_" + str(b) + ".xlsx"
dfsek4 = df[(df.SEKTION == "Sektion 4: Kulturwissenschaften")]
dfsek4 = dfsek4.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
path_sek4 = path + "/" + sektion4


# Sektion 5
b += 1
sektion5 = "sektion_" + str(b) + ".xlsx"
dfsek5 = df[(df.SEKTION == "Sektion 5: Wirtschafts- und Rechtswissenschaften")]
dfsek5 = dfsek5.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
path_sek5 = path + "/" + sektion5



# Sektion 6
b += 1
sektion6 = "sektion_" + str(b) + ".xlsx"
dfsek6 = df[(df.SEKTION == "Sektion 6: Gesellschaftswissenschaften")]
dfsek6 = dfsek6.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
path_sek6 = path + "/" + sektion6



# Sektion 7
b += 1
sektion7 = "sektion_" + str(b) + ".xlsx"
dfsek7 = df[(df.SEKTION == "Sektion 7: Wissenschaft – Technik – Gesellschaft")]
dfsek7 = dfsek7.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
path_sek7 = path + "/" + sektion7


# Kommissionen / Kuratorien
komm = "kommissionen_kuratorien.xlsx"
dfkomm = df[(df.SEKTION == "Kommission / Kuratorium")]
dfkomm = dfsek7.sort_values(by=["MITGLIEDINSTITUTION", "FORM_ID", "REFERENZ-NR."])
path_komm = path + "/" + komm

if os.path.exists(path_sek1):
    os.remove(path_sek1)
    print("The file " + sektion1 + " has been deleted successfully")
    dfsek1.to_excel(path + "/" + sektion1)
else:
    print("The file "  + sektion1 + " does not exist!")
    dfsek1.to_excel(path + "/" + sektion1)

if os.path.exists(path_sek2):
    os.remove(path_sek2)
    print("The file " + sektion2 + " has been deleted successfully")
    dfsek2.to_excel(path + "/" + sektion2)
else:
    print("The file "  + sektion2 + " does not exist!")
    dfsek2.to_excel(path + "/" + sektion2)

if os.path.exists(path_sek3):
    os.remove(path_sek3)
    print("The file " + sektion3 + " has been deleted successfully")
    dfsek3.to_excel(path + "/" + sektion3)
else:
    print("The file "  + sektion3 + " does not exist!")
    dfsek3.to_excel(path + "/" + sektion3)
    
if os.path.exists(path_sek4):
    os.remove(path_sek4)
    print("The file " + sektion4 + " has been deleted successfully")
    dfsek4.to_excel(path + "/" + sektion4)
else:
    print("The file "  + sektion4 + " does not exist!")
    dfsek4.to_excel(path + "/" + sektion4)
    
if os.path.exists(path_sek5):
    os.remove(path_sek5)
    print("The file " + sektion5 + " has been deleted successfully")
    dfsek5.to_excel(path + "/" + sektion5)
else:
    print("The file "  + sektion5 + " does not exist!")
    dfsek5.to_excel(path + "/" + sektion5)
    
if os.path.exists(path_sek6):
    os.remove(path_sek6)
    print("The file " + sektion6 + " has been deleted successfully")
    dfsek6.to_excel(path + "/" + sektion6)
else:
    print("The file "  + sektion6 + " does not exist!")
    dfsek6.to_excel(path + "/" + sektion6)
    
if os.path.exists(path_sek7):
    os.remove(path_sek7)
    print("The file " + sektion7 + " has been deleted successfully")
    dfsek7.to_excel(path + "/" + sektion7)
else:
    print("The file "  + sektion7 + " does not exist!")
    dfsek7.to_excel(path + "/" + sektion7)
    
if os.path.exists(path_komm):
    os.remove(path_komm)
    print("The file " + komm + " has been deleted successfully")
    dfkomm.to_excel(path + "/" + komm)
else:
    print("The file "  + komm + " does not exist!")
    dfkomm.to_excel(path + "/" + komm)


print()
print("Speichern abgeschlossen!")
print()
print("Herzlichen Dank für Ihren Besuch und bis bald!")
sys.exit(0)
