# -*- coding: utf-8 -*-
"""
Vers.1.5, 4.9.2020, Eberhard Waffenschmidt, TH-Köln
weitergeführt von Patrick Lehnen, TH-Köln

Tool zum Einlesen und Weiterverarbeiten von ILIAS Testergebnissen.
Liest die Ergebnisse eines ILIAS-Formelfragentest aus einer EXCEL-Tabelle ein.
Liest ebenso einen Fragenpool aus einer EXCEL-Tabelle ein.
Verknüpft beides und macht damit eine elektronische Korrektur des ILIAS-Tests 
anhand der im EXCEL-Fragenpool hinterlegten Formel.

Diese Version ist folgendermaßen limitiert:
- Auswertung von ausschließlich Formelfragen. 
Single- und Multiple-Choice und andere Fragetypen werden nicht unterstützt. 
Das Tool ist nicht getestet, was dann passiert.
- Fragen dürfen nur eine Antwort haben. 
Fragen mit mehr als einer Antwort (also z.B. mehreren Unterpunkten) werden 
nicht unterstützt. Es wird dann nur die erste Antwort ausgewertet.

Wichtige Parameter können (und müssen) als Konstanten diekt am Anfang des Codes 
festgelegt werden. Dazu gehören:
  Anzahl Fragen, 
  Anzahl Variablen pro frage,
  Anzahl Ergebnisse pro frage (hier fest auf 1, bitte nicht ändern)
  Max. Punktzahl im Test
  Notenschema
  Filename zum Excel-Export

ILIAS-Ergebnis-Datei:
Der titel muss lauten "ILIAS_Testergebnisse.xlsx".
Bitte die aus ILIAS exportierte Datei entsprechend umbenennen
Die Datei muss im selben Verzeichnis wie das Python-Skript liegen.
Erzeugung im ILIAS-Test: 
-> [Statistik] -> "Evaluationsdaten exportieren als" "Microsoft Exel" 
-> [Export]
Export dauert of recht lange.
Gebraucht und eingelesn wird das Datenblatt "Auswertung für alle Benutzer". 

Das Datenformat in der EXCEL-Tabelle sollte folgendermaßen aussehen:
---------------------------------------------------------------------
Ergebnisse von Testdurchlauf 1 für Max, Mustermann| 
            |	
Formelfrage	|26.1.1 Blindwiderstand einer Induktivität
$v1	        |83.5
$v2	        |46.1
$r1	        |132.916
            |	
Formelfrage	|09.1.1 maximale Leistung von Quelle
$v1	        |60
$v2	        |24
$r1	        |1440
:
usw.
:
Ergebnisse von Testdurchlauf 1 für Maria, Musterfrau| 
            |	
Formelfrage	|18.1.1 Induzierte Spannung 
$v1	        |340
$v2	        |3
$v3	        |35
$v4	        |60
---------------------------------------------------------------------

  Der name steht also immer in der ersten Spalte nach dem Schlüsseltext
"Ergebnisse von Testdurchlauf 1 für "
  Dann folgt eine Leerzeile (wird hier nicht ausgewertet).
  Dann folgen die einzelnen Fragen. Eine Formelfrage fängt mit dem 
Schlüsseltext "Formelfrage" in der ersten Spalte an. In der zweiten Spalte 
steht dann der titel der frage.
  Dieses Tool nimmt an, dass der Text bis zum ersten 
Leerzeichen der ID der frage entspricht und extrahiert diesen zusätzlich zum
titel.
  Nach dem titel folgen die Variablen mit den für den teilnehmer generierten
Variablenwerten. Es werden nur die Variablen aufgelistet, die auch in der 
frage verwendet wurden. Der ILIAS-name ($v1 usw.) steht in der ersten Spalte,
der dazugehörige Wert in der zweiten. ACHTUNG: Wenn der teilnehmer die frage
gar nicht geöffnet hat, werden keine Werte generiert und hier nicht aufgelistet.
  Dann werden die vom teilnehmer berechneten Werte $r1 usw. angezeigt. Auch hier
taucht $r1 nur auf, wenn der teilnehmer auch eine Eingabe gemacht hat.
ACHTUNG: Diese Version berücksichtigt nur das erste Ergebnis $r1.

Fragenpool-EXCEL-Tabelle
Der titel muss lauten "ILIAS_Fragenpool.xlsx".
Bitte die Datei entsprechend umbenennen.
Die Datei muss im selben Verzeichnis wie das Python-Skript liegen.
Die Fragenpool-Tabelle hat ein eigenes Format.
Jede zeile entspricht einer frage im Fragenpool
Die Tabelle muss die passenden Spaltenüberschriften haben. 
Diese müssen in der 6. zeile stehen, denn die ersten 5 Zeilen werden übersprungen.
Folgende Überschriften müssen mit den dazugehörigen Spalten vorhanden sein. 
Dabei ist die Reihenfolge der Spalten egal, es können auch noch andere dazwischen sein.
Beispiel:
----------------------------------------------------
: 5x Leerzeile
Question Title         |Formula 1|res1 tol|res1 pts
02.2.3 Ohmsches Gesetz |$v1/$v2  |5       |1
03.1.1 Leistung        |$v1*$v2  |5       |1
:
----------------------------------------------------    
Question Title: ist der titel der frage. Dieses Tool nimmt an, dass der Text
   bis zum ersten Leerzeichen der ID der frage entspricht und extrahiert
   diesen zusätzlich zum titel.
   Dieses Tool matcht Fragen aus der ILIAS-Ergebnisdatei mit dem Fragenpool
   anhand dieser Fragen-ID. 
Formula 1: Enthält die Formel zur Musterlösung im ILIAS Format. Nur die Formel 
   zu Antwort 1 wird ausgewertet.
res1 tol: Enthält die toleranz des ERgebnisses in %. Hier als +/-5%
res1 pts: Enthält die Anzahl punkte für die richtige Antwort.

Export-Excel-Datei:
Der name der Datei kann als Konstante am Anfang des Codes festgelegt werden.
Default ist "Testergebnisse.xlsx"
- Alle Ausgaben werden in ein Blatt geschrieben
- Alle Daten eines Teilnehmers entsprechen einer zeile
- Spaltenaufteilung:
    Nr name Vorname Familienname mat_nr Note GesPkt A1Pkt ... A40Pkt ...
    ... A1_ID A01_Formel A01_Tol A1_ResRef A1_Res  A1_v1 A1_v2...A1_v10 ...
    ...
    ... A40_ID A01_Formel A40_Tol A40_ResRef A40_Res  A40_v1 A40_v2...A40_v10

##############################################################################
# HISTORY ####################################################################
##############################################################################
Vers.1.5, 4.9.2020, E. Waffenschmidt:
    - Berücksichtigt ein Notenschema
    - Doku zu Dateiformaten
Vers.1.4, 4.9.2020, E. Waffenschmidt:     
  - Export der Ergebnisse nach EXCEL 
    - Alle Ausgaben werden in ein Blatt geschrieben
    - Alle Daten eines Teilnehmers entsprechen einer zeile
    - Spaltenaufteilung:
    Nr name Vorname Familienname mat_nr Note GesPkt A1Pkt ... A40Pkt ...
    ... A1_ID A01_Formel A01_Tol A1_ResRef A1_Res  A1_v1 A1_v2...A1_v10 ...
    ...
    ... A40_ID A01_Formel A40_Tol A40_ResRef A40_Res  A40_v1 A40_v2...A40_v10
  
Vers.1.3, 3.9.2020, E. Waffenschmidt: 
  - Verknüpfung von Testergebnissen mit Fragenpool
  - Funktionen zur Initialisierung von Arrays

Vers.1.2, 3.9.2020, E. Waffenschmidt: 
  - Im Prinzip unbegrenzte Anzahl von Variablen und Results pro frage.
    Maximale Anzahl durch Konstante in Header festgelegt.
    ACHTUNG: Die aktuelle Version kann nur die Formel für das 1. Ergebnis 
    aus dem Fragenpool laden. Für mehr Ergebnisse muss die 
    entsprechende Routine geändert werden.
  - Ermittlung der frage-ID aus dem titel als Funktion ausgelagert.
  - Einlesen der ILIAS-Fragen mit titel, Formel usw.
  - Funktion zur Evaluation der Formel.

Vers.1.1, 2.9.2020, E. Waffenschmidt: 
    Liest namen der teilnehmer, fragentitel der einzelnen teilnehmer sowie
    die Inhalte der ILIAS-Variablen zu den einzelnen Fragen ein. 
    Zusätzlich wird aus den Fragentiteln die fragen_id,
    also die Nummer der frage extrahiert.
    Weiterhin werden Vor- und Familienname extrahiert. 
    Die Annahme dabei ist, dass im namen erst der Familienname kommt,
    dann durch ", "(Komma und Spc) getrennt der Vorname.
    Die Ergebnisse sind in 1D bzw 2D Listen verfügbar.
"""

# Konstanten ###############################################################
anz_fragen = 14  # Anzahl Fragen pro teilnehmer
anz_var = 15     # Maximale Anzahl von frage-Variablen pro frage
anz_res = 5      # Maximale Anzahl von frage-Ergebnissen pro frage.
max_pkt = 58     # Maximale Anzahl von Punkten im Test
schema_note = ["5,0", "4,0", "3,7", "3,3", "3,0", "2,7", "2,3", "2,0", "1,7", "1,3", "1,0"]
# Mindestprozentzahl an Punkten für die korrespondierende Note
schema_proz = [0, 50, 54, 58, 62, 66, 70, 74, 78, 82, 86]
filename_export = "TestergebnisseGE2.xlsx"
ergebnis_datei = 'ILIAS_TestergebnisseGE2.xlsx'
fragen_datei = 'ILIAS_FragenpoolGE2.xlsx'
irtf_name = "irt_frame_ge2.xlsx"
############################################################################

import pandas as pd
import numpy as np
from math import *
from copy import deepcopy


def notenberechnung (pkt, max_pkt, schema_proz, schema_note):
    """ E.Waffenschmidt, 4.9.2020
    Berechnet für eine erzielte Punktzahl einer Prüfung 
    anhand eines Notenschemas die Note
    """
    note = "n.v."  # Default Note, wenn nichts gefunden wird. Eigentlich unmöglich.
    for i in range(len(schema_note)):
        if (pkt / max_pkt * 100) >= schema_proz[i]:
            note = schema_note[i]
    return note

def init_2d_none (m, n):
    """ E. Waffenschmidt, 3.9.2020
    initialisiert ein 2D-Array mit "None" als Inhalt
    Zugriff mit x = a[m][n]
    """
    a = []
    for i in range(m):
        a.append([None] * n)
    return a

def init_2d_no_str (m, n):
    """ E. Waffenschmidt, 3.9.2020
    initialisiert ein 2D-Array mit Leerstring als Inhalt
    Zugriff mit x = a[m][n]
    """
    a = []
    for i in range(m):
        a.append([""] * n)
    return a

def init_3d_none (m, n, o):
    """ E. Waffenschmidt, 3.9.2020
    initialisiert ein 3D-Array mit "None" als Inhalt
    Zugriff mit x = a[m][n][o]
    """
    a = []
    for k in range(m):
        a.append(init_2d_none(n, o))
    return a

def init_3d_no_str (m, n, o):
    """ E. Waffenschmidt, 3.9.2020
    initialisiert ein 3D-Array mit Leerstring als Inhalt
    Zugriff mit x = a[m][n][o]
    """
    a = []
    for k in range(m):
        a.append(init_2d_no_str(n, o))
    return a

#def Get_Frage_ID (titel):
#    """ E. Waffenschmidt, 3.9.2020
#    Extrahiert die Fragen-ID (Die "Nummer" der frage) aus dem gesamten titel.
#    Konkret sind das alle Zeichen bis zum ersten Leerzeichen.
#    """
#    return titel[0:titel.find(" ")] #Übernimmt den Text in titel bis zum ersten Leerzeichen

def finde_fragenindex (fragen_id, fragen_id_pool):
    """ E. Waffenschmidt, 3.9.2020
    Ermittelt den Index einer frage in einem Fragebpool anhand der Fragen-ID
    Wenn die gesuchet ID nicht im Pool ist, wird "None" zurück geliefert.
    """
    index = None
    fragen_anz = len(fragen_id_pool)
    for id_nr in range(0, fragen_anz):
        if fragen_id == fragen_id_pool[id_nr]:
            index = id_nr
    return index

def eval_ilias_single(gleichung_ilias, v, r):
    """E. Waffenschmidt, 3.9.2020
       Evaluiert (d.h. nutzt die Gleichung zur Berechnung) 
       eine Formel im ILIAS-Format 
       Die Variablenwerte werden in Form einer Liste in der Variablen v übergeben.
       Die Anzahl der Variablen in der Liste ist beliebig 
       und ergibt sich aus der Länge der List-Variablen v.   
       ACHTUNG: Die erste variable $v1 wird zu v[0].
       Die Funktion nutzt den eval() Befehl in Python. 
       Fehler bei der Berechnung werden abgefangen, damit das Programm nicht abbricht. 
       Bei einem Fehler in der Formel wird eine Meldung auf der Console 
       ausgegeben und das Ergebnis ist None.
       Sicherheits-Limitierungen von eval müssen noch eingebaut werden.
       Erste vergebliche dazu Versuche siehe "Lade Exceltab Aufgabenpool V1_2.py"
       Zum berechnen wird die Gleichung vom ILIAS-Format an das Format für Python angepasst.
       Dabei werden alle Großbuchstaben in Kleinbuchstaben umgewandelt,
       weil die Fuktionen sonst nicht richtig sind.
       Ersetzungstabelle:
       ILIAS        Python
       $v1          v[1] etc.
       ^            **
                    pi
                    e
                    sin
                    sinh
       arcsin       asin
       arcsinh      asinh
                    cos
                    cosh
       arccos       acos
       arccosh      acosh
                    tan
                    tanh
       arctan       atan
       arctanh      atanh
                    sqrt
                    abs
       ln           log
       log          log10

    ILIAS kennt folgende Begriffe bei der Definition von Formeln:
    "Erlaubt ist die Verwendung von bereits definierten Variablen ($v1 bis $vn), 
    von bereits definierten Ergebnissen (z.B. $r1), 
    das beliebige Klammern von Ausdrücken, 
    die mathematischen Operatoren + (Addition), - (Subtraktion),
    * (Multiplikation), / (Division), ^ (Potenzieren), 
    die Verwendung der Konstanten 'pi' für die Zahl Pi und 'e‘ für die Eulersche Zahl, 
    sowie die mathematischen Funktionen 
    'sin', 'sinh', 'arcsin', 'asin', 'arcsinh', 'asinh', 'cos', 'cosh', 
    'arccos', 'acos', 'arccosh', 'acosh', 'tan', 'tanh', 'arctan', 'atan', 
    'arctanh', 'atanh', 'sqrt', 'abs', 'ln', 'log'."
    """
    anz_var = len(v)  # Anzahl der möglichen variable $v1, $v2 usw. in der ILIAS-Formel
    # Gleichung für Python-Format anpassen:
    gleichung_py = gleichung_ilias.lower()
    # Variablenbezeichnung umändern:
    for i in reversed(range(1, anz_var + 1)):
        var_ilias = "$v" + str(i)
        var_py = "v[" + str(i - 1) + "]"
        gleichung_py = gleichung_py.replace(var_ilias, var_py)
    for j in reversed(range(1, anz_var + 1)):
        res_ilias = "$r" + str(j)
        res_py = "r[" + str(j - 1) + "]"
        gleichung_py = gleichung_py.replace(res_ilias, res_py)
    # Mathematische Funktionen anpassen:
    gleichung_py = gleichung_py.replace("^", "**")
    gleichung_py = gleichung_py.replace("arcsin", "asin")
    gleichung_py = gleichung_py.replace("arcsinh", "asinh")
    gleichung_py = gleichung_py.replace("arccos", "acos")
    gleichung_py = gleichung_py.replace("arccosh", "acosh")
    gleichung_py = gleichung_py.replace("arctan", "atan")
    gleichung_py = gleichung_py.replace("arctanh", "atanh")
    gleichung_py = gleichung_py.replace("ln", "log")
    gleichung_py = gleichung_py.replace("log", "log10")

    # Gleichung checken und berechnen
    try:  # Testet, ob eine Fehler auftritt
        eval(gleichung_py)
    except:  # Wenn ein Fehler auftritt, Fehlermeldung
        result = None
        print("Gleichung ", gleichung_ilias, " enthält einen Fehler:")
        print("Python-Format:", gleichung_py)
        print("Variablen:", v)
        print("Ergebnisse:", r)
    else:  # sonst is alles OK, und die Gleichung wird berechnet.
        result = eval(gleichung_py)
    return result


def eval_ilias_batch(gleichungen, variablen, res):
    # Für den Fall, dass eine Unteraufgabe eine Lösung aus einer späteren Unteraufgabe benötigt,
    # werden
    k = 0
    while k < anz_res + 1:
        for id, gleichung_ilias in enumerate(gleichungen):
            if not res[id] is None or gleichung_ilias == " ":
                continue

            anz_var = len(variablen)  # Anzahl der möglichen variable $v1, $v2 usw. in der ILIAS-Formel
            # Gleichung für Python-Format anpassen:
            gleichung_py = gleichung_ilias.lower()
            # Variablenbezeichnung umändern:
            for i in reversed(range(1, anz_var + 1)):
                var_ilias = "$v" + str(i)
                var_py = "v[" + str(i - 1) + "]"
                gleichung_py = gleichung_py.replace(var_ilias, var_py)
            for j in reversed(range(1, anz_var + 1)):
                res_ilias = "$r" + str(j)
                res_py = "r[" + str(j - 1) + "]"
                gleichung_py = gleichung_py.replace(res_ilias, res_py)
            # Mathematische Funktionen anpassen:
            gleichung_py = gleichung_py.replace("^", "**")
            gleichung_py = gleichung_py.replace("arcsin", "asin")
            gleichung_py = gleichung_py.replace("arcsinh", "asinh")
            gleichung_py = gleichung_py.replace("arccos", "acos")
            gleichung_py = gleichung_py.replace("arccosh", "acosh")
            gleichung_py = gleichung_py.replace("arctan", "atan")
            gleichung_py = gleichung_py.replace("arctanh", "atanh")
            gleichung_py = gleichung_py.replace("ln", "log")
            gleichung_py = gleichung_py.replace("log", "log10")

            # Gleichung checken und berechnen
            try:  # Testet, ob ein Fehler auftritt
                result = eval(gleichung_py)
            except:  # Wenn ein Fehler auftritt, Fehlermeldung
                result = None
                print("Gleichung ", gleichung_ilias, " enthält einen Fehler:")
                print("Python-Format:", gleichung_py)
                print("Variablen:", variablen)
                print("Ergebnisse:", res)
            r[id] = result
        k += 1
    return r


########################################################################
## HAUPTPROGRAMM ##############################################################
########################################################################

print("Tool zur externen Bewertung von ILIAS Formelfragen-Tests")
print("Version 1.5, 4.9.2020")
print("(c) by Eberhard Waffenschmidt, TH-Köln")

# weitere Konstanten
name_marker = "Ergebnisse von Testdurchlauf 1 für "
fragentitel_marker = "Formelfrage"

dummytext = "xyz"

# Daten aus EXCEL-File einlesen:
# Vor dem Filename in '' muss ein "r" gesetzt werden.
# Das Blatt im EXCEL-File wird nochmal mit sheet_name benannt
# header=None wird gesetzt, wenn keine Spaltenüberschriften existieren.
# skiprows=5 überspringt die ersten 5 Zeilen 

# Testergebnisse einlesen:
#df1 = pd.read_excel (r'C:\Users\Ebi\Documents\python\ILIAS Prüfungs-Auswertung\ETAT-Probe-Klausur_results kurz.xlsx', sheet_name='Auswertung für alle Benutzer')
#df1 = pd.read_excel (r'C:\Users\Ebi\Documents\python\ILIAS Prüfungs-Auswertung\ETAT-Probe-Klausur_results mit Titelzeile.xlsx', sheet_name='Auswertung für alle Benutzer', header=None)

# Manche ILIAS-Exporte haben je einen Reiter pro Student. Die folgende Funktion überprüft
# die Beschaffenheit der Tabelle und formt sie entsprechend um

try:
    df1 = pd.read_excel(ergebnis_datei, sheet_name='Auswertung für alle Benutzer', header=None)
except:
    df_dict = pd.read_excel(ergebnis_datei, sheet_name=None, header=None)
    df1 = df_dict[list(df_dict.keys())[1]]
    for id in list(df_dict.keys())[2:]:
        new_student = df_dict[id]
        df1 = pd.concat([df1, new_student], ignore_index=True)

# Fragenpool aus Excel-Tabelle einlesen
FRAGEN_AUS_ILIAS_EXPORT = True
if FRAGEN_AUS_ILIAS_EXPORT:
    df2 = pd.read_excel(fragen_datei, sheet_name='SQL - Database')
else:
    df2 = pd.read_excel(fragen_datei, sheet_name='Tabelle1', skiprows=5)
#df2 = pd.read_excel (r'C:\Users\Ebi\Documents\python\ILIAS Prüfungs-Auswertung\EGT-ILIAS-Klausuraufgaben  edEW-FK4_1.xlsx', sheet_name='Tabelle1', skiprows=5)
########################################################################
### Testergebnisse verarbeiten: ########################################
########################################################################
# Daten in 2D-Array umwandeln, damit der Zugriff einfacher zu indexieren ist
i_d = df1.values #i_d steht für ILIAS-Daten
anz_zeilen = len(i_d)

# Anzahl teilnehmer ermitteln
anz_teilnehmer = 0
for zeile in range(anz_zeilen):
   txt = i_d[zeile, 0]
   if (txt.__class__ == dummytext.__class__):
      if txt.startswith(name_marker):
         anz_teilnehmer = anz_teilnehmer + 1

# variable initialisieren
nr_teilnehmer = [None] * anz_teilnehmer # Fortlaufende Nummer
namen = [""] * anz_teilnehmer
vornamen = [""] * anz_teilnehmer
familiennamen = [""] * anz_teilnehmer
mat_nr = [""] * anz_teilnehmer
noten = [""] * anz_teilnehmer
fragentitel = init_3d_no_str(anz_teilnehmer, anz_fragen, anz_res)  # fragentitel = [[""]*anz_fragen]*anz_teilnehmer fuktioniert nicht!
fragen_id = init_3d_no_str(anz_teilnehmer, anz_fragen, anz_res)
fragen_formel = init_3d_no_str(anz_teilnehmer, anz_fragen, anz_res)
fragen_tol = init_3d_none(anz_teilnehmer, anz_fragen, anz_res)
var = init_3d_none(anz_teilnehmer, anz_fragen, anz_var)  # Zugriff mit: V = var[teilnehmer][frage][variable]
res = init_3d_none(anz_teilnehmer, anz_fragen, anz_res)
res_ref = init_3d_none(anz_teilnehmer, anz_fragen, anz_res)  # Richtiges Ergebnis als Referenz
pkt = init_3d_none(anz_teilnehmer, anz_fragen, anz_res)     # Vergebene punkte für die Aufgabe
ges_pkt = [None] * anz_teilnehmer

# Daten der teilnehmer analysieren
print("Testergebnisse werden analysiert...")
teilnehmer = 0
frage_nr = 0
for zeile in range(anz_zeilen):
   txt = i_d[zeile, 0]
   if (txt.__class__ == dummytext.__class__):
      if txt.startswith(name_marker):
         teilnehmer = teilnehmer + 1
         nr_teilnehmer[teilnehmer - 1] = teilnehmer
         name = txt.replace(name_marker, "")  # Der Text vor dem namen wird entfernt
         namen[teilnehmer - 1] = name
         familiennamen[teilnehmer - 1] = name[0:name.find(",")] #Übernimmt den Text in titel bis zum ersten Komma.
         vornamen[teilnehmer - 1] = name[name.find(",") + 2:]     #Übernimmt den Text in titel 2 Stellen nach dem ersten Komma.
         # print ("teilnehmer Nr.: ",teilnehmer, " : ",name," Vorname: ",vornamen[teilnehmer-1], " Familienname: ",familiennamen[teilnehmer-1] )
         frage_nr = 0
      if txt.startswith(fragentitel_marker):
         frage_nr = frage_nr + 1
         titel = i_d[zeile, 1]
         fragentitel[teilnehmer - 1][frage_nr - 1] = titel
         fragen_id[teilnehmer - 1][frage_nr - 1] = titel  # Get_Frage_ID (titel)
#         print ("teilnehmer Nr.: ",teilnehmer, " Fragenummer: ",frage_nr, " ID: ",fragen_id[teilnehmer-1][frage_nr-1])
      for var_nr in range(1, anz_var + 1):
          var_marker = "$v" + str(var_nr)
          if txt == var_marker:
             x = i_d[zeile, 1]
             var[teilnehmer - 1][frage_nr - 1][var_nr - 1] = x
#             print ("var[",var_nr,"] = ",x)
      for res_nr in range(1, anz_res + 1):
          res_marker = "$r" + str(res_nr)
          if txt.startswith(res_marker):
             y = i_d[zeile, 1]
             res[teilnehmer - 1][frage_nr - 1][res_nr - 1] = y
#             print ("res[",res_nr,"] = ",y," - ",res[teilnehmer-1][frage_nr-1], "Vorh.r=",res[teilnehmer-1][frage_nr-2],)
      
########################################################################
### Fragenpool verarbeiten #############################################
########################################################################
# relevante Spalten in 1D-Arrays (Lists) kopieren
if FRAGEN_AUS_ILIAS_EXPORT:
    titels_pool = df2['question_title']
    gleichungen_pool = {x: df2[f'res{x + 1}_formula'] for x in range(anz_res)}
    toleranzen_pool = {x: df2[f'res{x + 1}_tol'] for x in range(anz_res)}
    punkte_pool = {x: df2[f'res{x + 1}_points'] for x in range(anz_res)}
else:
    titels_pool = df2['Question Title']
    gleichungen1_pool = df2['Formula 1']
    toleranzen1_pool = df2['res1 tol']
    punkte1_pool = df2['res1 pts']
 # Nur Formel 1 wird ausgewertet

anz_fragen_pool = len(titels_pool)

#fragen_id_pool = [""] * anz_fragen_pool
# Fragen IDs aus den Titeln extrahieren:
#for i in range(0, anz_fragen_pool):
#    fragen_id_pool[i] = Get_Frage_ID(titels_pool[i])

fragen_id_pool = list(titels_pool)
    
########################################################################
### Testergebnisse und Fragenpool verknüpfen ###########################
########################################################################

# Formeln zu den Ergebnissen der teilnehmer zuordnen,
# Richtiges Ergebnis berechnen 
# Mit den Ergebnis des Studenten vergleichen 
# und punkte vergeben
print("Testergebnisse werden mit Fragenpool verknüpft...")
for std_id, std_aufgaben in enumerate(fragen_id):
    ges_punkte = 0
    for afg_id, aufgabe in enumerate(std_aufgaben):
        fid = finde_fragenindex(fragen_id[std_id][afg_id], fragen_id_pool)
        fragen_punkte_student = [0 for _ in range(anz_res)]
        musterloesung = [None for _ in range(anz_res)]
        formeln = [gleichungen_pool[x][fid] for x in range(anz_res)]
        toleranzen = [toleranzen_pool[x][fid] for x in range(anz_res)]
        punkte = [punkte_pool[x][fid] for x in range(anz_res)]
        std_variablen = var[std_id][afg_id]
        std_ergebnisse = res[std_id][afg_id]
        if not std_variablen[0] == None:  # Wenn der Student die Aufgabe gar nicht angeschaut hat, sind alle Variablen None, insbesondere die erste.
            musterloesung = eval_ilias_batch(formeln, std_variablen, musterloesung)
            # Maximale und minimale Grenze mit toleranz bestimmen.
            # ACHTUNG: Bei negativem Vorzeichen drehen sich min und max rum,
            # das gibt dann Ärger beim nachfolgenden Vergleich
            # Daher hier die Verwendung von "min" und "max"
            for uid, (r_student, r, toleranz, punkte) in enumerate(zip(std_ergebnisse,
                                                                       musterloesung,
                                                                       toleranzen,
                                                                       punkte)):
                if r_student is not None:  # Wenn der Student keine Lösung angegeben hat ist r_student = None
                    # Dann kann es noch sein, dass die Lösung als Bruch, z.B. 1/300 angegebn ist.
                    # Dann muss der Bruch mit eval ausgerechnet werden
                    if r is None:  # Überprüfung ob Berechnung eines Ergebnisses nicht geklappt hat
                        print(f"Fehler in Berechnung der Formel: {aufgabe}\nUnteraufgabe: {uid + 1}")
                    r_min = min(r * (1 + toleranz / 100), r * (1 - toleranz / 100))
                    r_max = max(r * (1 + toleranz / 100), r * (1 - toleranz / 100))
                    if type(r_student) == str:
                        r_student = eval(r_student)
                    if (r_student >= r_min) and (r_student <= r_max):
                        fragen_punkte_student[uid] = punkte
                        ges_punkte = ges_punkte + punkte
                    for i, gleichung in enumerate(formeln): #Überprufung ob es überhaupt eine Fragen Unterteil gab
                        if gleichung == " ":
                            fragen_punkte_student[i] = None

            fragen_formel[std_id][afg_id] = formeln
            fragen_tol[std_id][afg_id] = toleranzen
            res_ref[std_id][afg_id] = musterloesung
            pkt[std_id][afg_id] = fragen_punkte_student
    ges_pkt[std_id] = ges_punkte
    noten[std_id] = notenberechnung(ges_punkte, max_pkt, schema_proz, schema_note)


"""
for teilnehmer in range(0,anz_teilnehmer):
    # print ("Nr.",teilnehmer,namen[teilnehmer],":")
    ges_punkte = 0
    for frage in range(0,anz_fragen):
#        print("Teiln.",teilnehmer,"frage:",frage,"ID:",fragen_id[teilnehmer][frage])
#        print("Variablen:",var[teilnehmer][frage])
#        print("Stud.res.:",res[teilnehmer][frage])

        # print("")
        fragen_punkte_student = 0
        # print ("Teiln.",teilnehmer,", frage",frage,",",fragen_id[teilnehmer][frage])
        # Index der frage im Fragenpool finden
        i = finde_fragenindex(fragen_id[teilnehmer][frage], fragen_id_pool)
        Musterlösung = [None for _ in range(anz_res)]
        for res_range in range(anz_res):

            print(fragen_id_pool[i])

            if i==None: #Falls frage nicht im Pool gefunden wird
                r = -999999
                Formel = "Formel Nicht gefunden"
                toleranz = None
                print("!!! Teiln.",teilnehmer,", frage",frage,",",fragen_id[teilnehmer][frage]," existiert nicht im Pool!")
            else:
                # print ("Fragenindex i =",i)
                # Passende Formel usw. aus dem Fragenpool auslesen
                # Derzeit wird nur Ergebnis 1 ausgewertet
                Formel = gleichungen_pool[res_range][i]
                toleranz = toleranzen_pool[res_range][i]
                punkte = punkte_pool[res_range][i]
                if Formel == " " or Formel is np.nan:
                    continue
                # Formel mit den Variablen des Studenten anwenden:
                v = var[teilnehmer][frage]

                # Ergebnisse der Studierenden für Folgefehler-Berechnung auskommentieren
                # r = res[teilnehmer][frage]
                if v[0] == None: # Wenn der Student die Aufgabe gar nicht angeschaut hat, sind alle Variablen None, insbesondere die erste.
                    r = None
                else:  # Der Student hat wenigstens die frage angeschaut und Werte bekommen
                    r = eval_ilias_single(Formel, v, Musterlösung)
                    if r is None:
                        apa = {"T": toleranz}
                    Musterlösung[res_range] = r
                    # print ("Formel =",Formel,"=",r,"toleranz:",toleranz,"%")
                    # print ("Typ von r = ",type(r))

                    # Maximale und minimale Grenze mit toleranz bestimmen.
                    # ACHTUNG: Bei negativem Vorzeichen drehen sich min und max rum,
                    # das gibt dann Ärger beim nachfolgenden Vergleich
                    # Daher hier die Verwendung von "min" und "max"
                    r_min = min(r*(1+toleranz/100), r*(1-toleranz/100))
                    r_max = max(r*(1+toleranz/100), r*(1-toleranz/100))
                    # print("Min =",r_min,", Max =",r_max)

                    # Ergebnis des Studierenden:
                    r_student = res[teilnehmer][frage][res_range]  # Es wird in dieser version nur ein Ergebnis, das erste, ausgewertet
                    # print("Stud. Ergebnis: Teiln.",teilnehmer,"frage:",frage,"R_stud =",r_student)

                    # Ist das Ergebnis vorhanden und innerhalb der toleranz?
                    # Dann gibt's die punkte für die Aufgabe, sonst 0 pkt.
                    if r_student != None: #Wenn der Student keine Lösung angegeben hat ist r_student = None
                        # Dann kann es noch sein, dass die Lösung als Bruch, z.B. 1/300 angegebn ist.
                        # Dann muss der Bruch mit eval ausgerechnet werden
                        if type(r_student)==str:
                            r_student = eval(r_student)
                        if (r_student >= r_min) and (r_student <= r_max):
                            fragen_punkte_student += punkte
                            ges_punkte = ges_punkte + punkte
                            # print("punkte =",Punkte_Student)
                
        # Jetzt noch die Ergebnisse in den Listen abspeichern:
        # print ("frage",frage,",",fragen_id[teilnehmer][frage],"pkt =",fragen_punkte_student)
        fragen_formel[teilnehmer][frage] = Formel
        fragen_tol[teilnehmer][frage] = toleranz
        res_ref[teilnehmer][frage][0] = r
        pkt[teilnehmer][frage][0] = fragen_punkte_student
    ges_pkt[teilnehmer] = ges_punkte
    noten[teilnehmer] = notenberechnung (ges_punkte, max_pkt, schema_proz, schema_note)
    print ("Nr.",nr_teilnehmer[teilnehmer],namen[teilnehmer],", Ges.pkt =",ges_pkt[teilnehmer],", Note =",noten[teilnehmer])
"""
print("Anzahl Teilnehmer = ", anz_teilnehmer)

########################################################################
### Daten in EXCEL-Sheet exportieren  ##################################
########################################################################
# Dazu passenden Pandas-Dataframe zusammenbauen:
# Zeilentitel generieren:
#  Nr name Vorname Familienname mat_nr Note GesPkt A1Pkt ... A40Pkt ...
#  ... A1_ID A01_Formel A01_Tol A1_ResRef A1_Res  A1_v1 A1_v2...A1_v10 ...
#  ...
#  ... A40_ID A01_Formel A40_Tol A40_ResRef A40_Res  A40_v1 A40_v2...A40_v10
    
print('Daten werden nach EXCEL exportiert...')

df_ex = pd.DataFrame() # leeren DataFrame zum Export erzeugen
# Übersichtsdaten
df_ex['Nr'] = nr_teilnehmer  # erzeugt eine neue Spalte mit dem titel 'Nr' und Daten in nr_teilnehmer
df_ex['name'] = namen  # erzeugt eine neue Spalte mit dem titel 'name' und den Daten in namen
df_ex['Vorname'] = vornamen
df_ex['Familienname'] = familiennamen
df_ex['mat_nr'] = mat_nr
df_ex['Note'] = noten
df_ex['GesPkt'] = ges_pkt

# Gesamtpunkte bei den einzelnen Fragen
for frage in range(anz_fragen):
    for untertitel in range(anz_res):
        spaltentitel = "A" + str(frage + 1) + f".{untertitel + 1}_Pkt"
        x = [None] * anz_teilnehmer  # Spaltendaten initialisieren
        for teilnehmer in range(anz_teilnehmer):
            x[teilnehmer] = pkt[teilnehmer][frage][untertitel]
        df_ex[spaltentitel] = x

for frage in range(anz_fragen):
    spaltentitel = "A" + str(frage + 1) + f".{untertitel + 1}_Pkt_Gesamt"
    x = [None] * anz_teilnehmer  # Spaltendaten initialisieren
    for teilnehmer in range(anz_teilnehmer):
        x[teilnehmer] = sum(filter(None, pkt[teilnehmer][frage]))
    df_ex[spaltentitel] = x

df_ex[''] = [""] * anz_teilnehmer  # Leerspalte an dieser Stelle einfügen

# Details zu den einzelnen Fragen 
for frage in range (anz_fragen):
    spaltentitel = f"A" + str(frage + 1) + f"_ID"
    x = [""] * anz_teilnehmer  # Spaltendaten initialisieren
    for teilnehmer in range(anz_teilnehmer):
        x[teilnehmer] = fragen_id[teilnehmer][frage]
    df_ex[spaltentitel] = x
    for untertitel in range(anz_res):
        spaltentitel = "A" + str(frage + 1) + f".{untertitel}_Formel"
        x = [""] * anz_teilnehmer  # Spaltendaten initialisieren
        for teilnehmer in range(anz_teilnehmer):
            x[teilnehmer] = fragen_formel[teilnehmer][frage][untertitel]
        df_ex[spaltentitel] = x

    spaltentitel = "A" + str(frage + 1) + "_Tol"
    x = [""] * anz_teilnehmer  # Spaltendaten initialisieren
    for teilnehmer in range(anz_teilnehmer):
        x[teilnehmer] = fragen_tol[teilnehmer][frage]
    df_ex[spaltentitel] = x

    for untertitel in range(anz_res):
        spaltentitel = "A" + str(frage + 1) + f".{untertitel}_Res_Ref"
        x = [""] * anz_teilnehmer  # Spaltendaten initialisieren
        for teilnehmer in range(anz_teilnehmer):
            x[teilnehmer] = res_ref[teilnehmer][frage][untertitel]
        df_ex[spaltentitel] = x

    for untertitel in range(anz_res):
        spaltentitel = "A" + str(frage + 1) + f".{untertitel}_Res"
        x = [""] * anz_teilnehmer # Spaltendaten initialisieren
        for teilnehmer in range(anz_teilnehmer):
            x[teilnehmer] = res[teilnehmer][frage][untertitel]
        df_ex[spaltentitel] = x
    
    # Variablen der einzelnen Fragen
    for variable in range(anz_var):
        spaltentitel = "A" + str(frage + 1) + "_v" + str(variable + 1)
        x = [""] * anz_teilnehmer  # Spaltendaten initialisieren
        for teilnehmer in range(anz_teilnehmer):
            x[teilnehmer] = var[teilnehmer][frage][variable]
        df_ex[spaltentitel] = x
    
# Datenframe in EXCEL-File schreiben   
df_ex.to_excel(filename_export, index=False) # Index = False sorgt dafür, dass die erste Spalte nicht den Zeilenidex von 0..Ende enthält

print('Fertig!')

# Hier folgen die Vorbereitungen für ein DataFrame, was mithilfe des IRT-Tools ausgewertet werden kann
#res_dict = {res_id: None for res_id in range(anz_res)}
#FragenID_dict = {fid: deepcopy(res_dict) for fid in fragen_id_pool}  # Jede Fragen ID erhält eine Spalte im späteren DataFrame
#irtf = {uid: deepcopy(FragenID_dict) for uid in range(anz_teilnehmer)}  # Jedem teilnehmer wird jede frage zugeordnet

#Im folgenden werden die punkte den entsprechenden Fragen pro teilnehmer zugeordnet
pkt_per_frage = {}
for uid, fids, pkte in zip(range(anz_teilnehmer), fragen_id, pkt):
    pkt_per_frage[uid] = {}

    for fid, pkt in zip(fids, pkte):
        for pid, p in enumerate(pkt):
            if p == 0:  # Abfrage ob Frage richtig berechnet wurde. Für IRT wäre -1: "Falsche Antwort" und 0: "frage nicht gestellt"
                p = -1
            if p is None:
                p = 0
            pkt_per_frage[uid][f"{fid}.{pid}"] = p

irtf = {}
for uid, frg_pool in enumerate(pkt_per_frage.values()):
    irtf[uid] = {}
    for frg in frg_pool:
        irtf[uid][frg] = frg_pool[frg]
        #print(uid, frg)
        #print(irtf[uid])

pd.DataFrame(irtf).T.fillna(0).to_excel(irtf_name)
# Leider müssen jetzt noch von Hand die Formatierung und die erste sowie letzte Spalte aus der Excel gelöscht werden