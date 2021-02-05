# -*- coding: utf-8 -*-
"""
Vers.1.13, 30.10.2020, Eberhard Waffenschmidt, TH-Köln

Tool zum Einlesen und Weiterverarbeiten von ILIAS Testergebnissen.
Liest die Ergebnisse eines ILIAS-Formelfragentest aus einer EXCEL-Tabelle ein.
Liest ebenso einen Fragenpool aus einer EXCEL-Tabelle ein.
Verknüpft beides und macht damit eine elektronische Korrektur des ILIAS-Tests 
anhand der im EXCEL-Fragenpool hinterlegten Formel.
Liest weiterhin eine Tabelle mit Bonuspunkten ein und verknüpft sie mit den 
Teilnehmern
Außerdem werden die Matrikelnummern aus der PSSO-Anmelde-Liste eingelesen. 
Nicht angemeldete Teilnehmer werden in der Console angezeigt 
und haben in der Ergebnisliste keine Matrikelnummer.

Diese Version ist folgendermaßen limitiert:
- Auswertung von ausschließlich Formelfragen und Freitextaufgabe
- Single- und Multiple-Choice und andere Fragetypen werden nicht unterstützt. 
  Diese werden ignoriert
- Fragen dürfen nur eine Antwort haben. 
  Fragen mit mehr als einer Antwort (also z.B. mehreren Unterpunkten) werden 
  nicht unterstützt. Es wird dann nur die erste Antwort ausgewertet.
- bei Freitextaufgaben können nur Fragen berücksichtigt werden,
  bei denen ein einzelnes Stichwort im Freitext enthalten sein muss.
  Das Stichwort für Freitextaufgaben muss im EXCEL_Fragenpool in der Spalte 
  für die Formel des ersten Ergebnisses stehen ("res1_formula").

Wichtige Parameter können (und müssen) als Konstanten diekt am Anfang des Codes 
festgelegt werden. Dazu gehören:
  Anzahl Fragen, 
  Anzahl Variablen pro Frage, 
  Anzahl Ergebnisse pro Frage (hier fest auf 1, bitte nicht ändern) 
  Max. Punktzahl im Test
  Notenschema
  Filename zum Excel-Export

ILIAS-Ergebnis-Datei:
Der Titel muss lauten "ILIAS_Testergebnisse.xlsx". 
Bitte die aus ILIAS exportierte Datei entsprechend kopieren und umbenennen.
Die Datei (wie alle anderen auch) muss im selben Verzeichnis 
wie das Python-Skript liegen.
Erzeugung im ILIAS-Test: 
-> [Statistik] -> "Evaluationsdaten exportieren als" "Microsoft Exel" 
-> [Export]
Export dauert of recht lange.
Eventuell bei dem Reiter "Export" nachschauen. manchmal findet sich dort eine 
passende exportierte Datei.

Gebraucht und eingelesen wird das Datenblatt "Auswertung für alle Benutzer". 

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

Freitext eingeben |50.1.01 Freitextfrage für ETAT
Ergebnis	      |Hier findet sich der Freitext vom Studi

---------------------------------------------------------------------

  Der Name steht also immer in der ersten Spalte nach dem Schlüsseltext 
"Ergebnisse von Testdurchlauf 1 für "
  Dann folgt eine Leerzeile (wird hier nicht ausgewertet).
  Dann folgen die einzelnen Fragen. Eine Formelfrage fängt mit dem 
Schlüsseltext "Formelfrage" in der ersten Spalte an. In der zweiten Spalte 
steht dann der Titel der Frage. 
  Dieses Tool nimmt an, dass der Text bis zum ersten 
Leerzeichen der ID der Frage entspricht und extrahiert diesen zusätzlich zum 
Titel.
  Nach dem Titel folgen die Variablen mit den für den Teilnehmer generierten 
Variablenwerten. Es werden nur die Variablen aufgelistet, die auch in der 
Frage verwendet wurden. Der ILIAS-Name ($v1 usw.) steht in der ersten Spalte,
der dazugehörige Wert in der zweiten. ACHTUNG: Wenn der Teilnehmer die Frage 
gar nicht geöffnet hat, werden keine Werte generiert und hier nicht aufgelistet.
  Dann werden die vom Teilnehmer berechneten Werte $r1 usw. angezeigt. Auch hier 
taucht $r1 nur auf, wenn der Teilnehmer auch eine Eingabe gemacht hat.
ACHTUNG: Diese Version berücksichtigt nur das erste Ergebnis $r1.

Fragenpool-EXCEL-Tabelle
Der Titel muss lauten "ILIAS_Fragenpool.xlsx". 
Bitte die Datei entsprechend umbenennen.
Die Datei muss im selben Verzeichnis wie das Python-Skript liegen.
Die Fragenpool-Tabelle hat ein eigenes Format.
Jede Zeile entspricht einer Frage im Fragenpool
Die Tabelle muss die passenden Spaltenüberschriften haben. 
Diese müssen in der 7. Zeile stehen, denn die ersten 6 Zeilen werden übersprungen.
Folgende Überschriften müssen mit den dazugehörigen Spalten vorhanden sein. 
Dabei ist die Reihenfolge der Spalten egal, es können auch noch andere dazwischen sein.
Beispiel:
----------------------------------------------------
: 6x Leerzeile
Question Type |Question Title          |res1_formula|res1 tol|res1 pts
Formelfrage   |02.2.3 Ohmsches Gesetz  |$v1/$v2     |5       |1
Formelfrage   |03.1.1 Leistung         |$v1*$v2     |5       |1
Freitextfrage |50.1.1 Farbe des Himmels|blau        |        |3
:
----------------------------------------------------    
Question Type : Beschreibt die Art der Frage. Kann entweder "Formelfrage" oder
   "Freitextfrage" haben.
Question Title: ist der Titel der Frage. Dieses Tool nimmt an, dass der Text 
   bis zum ersten Leerzeichen der ID der Frage entspricht und extrahiert 
   diesen zusätzlich zum Titel.
   Dieses Tool matcht Fragen aus der ILIAS-Ergebnisdatei mit dem Fragenpool
   anhand dieser Fragen-ID. 
res1_formula: Enthält die Formel zur Musterlösung im ILIAS Format. Nur die Formel 
   zu Antwort 1 wird ausgewertet. 
   Bei Freitextaufgaben steht hier das Schlüsselwort, das in der Antwort 
   enthalten sein muss
res1 tol: Enthält die Toleranz des Ergebnisses in %. Hier als +/-5%
res1 pts: Enthält die Anzahl Punkte für die richtige Antwort.

Bonuspunkte-EXCEL-Tabelle:
Hat den Namen "Bonuspunkte.xlsx"
Datenformat
----------------------------------------------------
Benutzername |Name                 |Bonuspunkte
sstudent     |Student, Sven        |2
sstuden2	 |Studentin, Svenja    |2
:
----------------------------------------------------    

PSSO-Tabelle
Entspricht der aus dem PSSO exportierten Teilnehmerliste.    
Hat den Namen "PSSO_Teilnehmer.xls"
Bitte entsprechend kopieren und umbennen.
Datenformat: 
----------------------------------------------------
: 3x Ignorierte Zeilen
:    
    mtknr |sortname         |nachname |vorname	|...
12345678  |Student,Sven     |Student  |Sven     |...
13456789  |Studentin,Svenja |Studentin|Svenja   |...
:    
----------------------------------------------------
sortname wird nicht verwendet, da der bei langen Namen verkrüppelt wird.
Stattdessen wird der Name aus "nachname und "vorname" zusammengesetzt

Export-Excel-Datei:
Der Name der Datei kann als Konstante am Anfang des Codes festgelegt werden.
Default ist "Python_Testergebnisse.xlsx"
- Alle Ausgaben werden in ein Blatt geschrieben
- Alle Daten eines Teilnehmers entsprechen einer Zeile
- Spaltenaufteilung:
    Nr Name Vorname Familienname MatNr Note GesPkt BonusPkt A1Pkt ... A40Pkt ...
    ... A1_ID A01_Formel A01_Tol A1_ResRef A1_Res  A1_v1 A1_v2...A1_v10 ...
    ...
    ... A40_ID A01_Formel A40_Tol A40_ResRef A40_Res  A40_v1 A40_v2...A40_v10

##############################################################################
# HISTORY ####################################################################
##############################################################################
Vers.1.13, 30.10.2020, Eberhard Waffenschmidt, TH-Köln
    - Aufgabe '22.1.03' wird rausgefilter und besonders behandelt:
      Bei der Lösung kann ein Winkel >360 rauskommen, da die Ranges schlecht 
      gewählt sind. Es werden nun Studi-Lösung und Python-Lösung jeweils
      in den Bereich 0..360° gebracht und dann erst verglichen. 
    - Export der finalen Noten in PSSO-Datei
Vers.1.12, 27.10.2020, Eberhard Waffenschmidt, TH-Köln
    - Exportiert öffentliche Ergebnisliste zur Einsicht als EXCEL-Tabelle
    - Buggy-Aufgabe wird manuell rausgefiltert:
      Aufgabe '07.1.04' ist buggy. Die Frage enthält Serienschaltung. 
      Gesamtspannung v2 in mV ist manchmal kleiner als Einzelspannung v3in V.
      Das ist physikalisch nicht möglich, 
      Die Lösung mit der angegebenen Formel ist dann Quatsch.
      Daher wird falls v2/1000 < v3 ist, ein Fehler erkannt
      und immer ein Gnadenpunkt vergeben. 
      und erst gar nicht die Formel berechnet. 
      Stattdessen ist das Ergebnis 0.
Vers.1.11, 18.9.2020, Eberhard Waffenschmidt, TH-Köln
    - Liest PSSO-Teilnehmer ein und ermittelt damit die MatNr.
    - MatNr aus Freitextaufgabe wird jetzt wieder ignoriert, da unzuverlässig.
Vers.1.10, 18.9.2020, Eberhard Waffenschmidt, TH-Köln
    - Liest Bonuspunkte ein und berücksichtigt sie in der Gesamtwertung
    - Übernimmt MatNr. aus Freitextaufgabe
Vers.1.9, 18.9.2020, Eberhard Waffenschmidt, TH-Köln
    - Bei Aufgaben, bei denen die Formel fehlerhaft ist,
      wird die volle Punktzahl vergeben. 
Vers.1.8, 17.9.2020, Eberhard Waffenschmidt, TH-Köln
    - Freitext-Aufgaben werden bei der Analyse der ILIAS-Ergebnisse 
      berücksichtigt. 
    - bei Freitextaufgaben können nur Fragen berücksichtigt werden,
      bei denen ein einzelnes Stichwort im Freitext enthalten sein muss.
      Das Stichwort muss im EXCEL_Fragenpool in der Spalte für die Formel 
      des ersten Ergebnisses stehen ("res1_formula").
    - "Question Type" aus Frage-Pool wird mitverwendet.
    - Versionsnummer als Konstante für Print
Vers.1.7, 16.9.2020, Eberhard Waffenschmidt, TH-Köln
    - Anzahl Fragen im Pool werden geprintet
Vers.1.6, 15.9.2020, Eberhard Waffenschmidt, TH-Köln
    - Print-Meldungen beim Dateneinlesen.
    - Fragenpool-Format an aktuelle Excel-Datei angepasst.
    - "," als Dezimalpunkt in einer Formel durch "." ersetzt
    - Wenn fehlerhafte Formel, dann ohne Ergebnis weiter
Vers.1.5, 4.9.2020, E. Waffenschmidt:
    - Berücksichtigt ein Notenschema
    - Doku zu Dateiformaten
Vers.1.4, 4.9.2020, E. Waffenschmidt:     
  - Export der Ergebnisse nach EXCEL 
    - Alle Ausgaben werden in ein Blatt geschrieben
    - Alle Daten eines Teilnehmers entsprechen einer Zeile
    - Spaltenaufteilung:
    Nr Name Vorname Familienname MatNr Note GesPkt A1Pkt ... A40Pkt ...
    ... A1_ID A01_Formel A01_Tol A1_ResRef A1_Res  A1_v1 A1_v2...A1_v10 ...
    ...
    ... A40_ID A01_Formel A40_Tol A40_ResRef A40_Res  A40_v1 A40_v2...A40_v10
  
Vers.1.3, 3.9.2020, E. Waffenschmidt: 
  - Verknüpfung von Testergebnissen mit Fragenpool
  - Funktionen zur Initialisierung von Arrays

Vers.1.2, 3.9.2020, E. Waffenschmidt: 
  - Im Prinzip unbegrenzte Anzahl von Variablen und Results pro Frage.
    Maximale Anzahl durch Konstante in Header festgelegt.
    ACHTUNG: Die aktuelle Version kann nur die Formel für das 1. Ergebnis 
    aus dem Fragenpool laden. Für mehr Ergebnisse muss die 
    entsprechende Routine geändert werden.
  - Ermittlung der Frage-ID aus dem Titel als Funktion ausgelagert. 
  - Einlesen der ILIAS-Fragen mit Titel, Formel usw.
  - Funktion zur Evaluation der Formel.

Vers.1.1, 2.9.2020, E. Waffenschmidt: 
    Liest Namen der Teilnehmer, Fragentitel der einzelnen Teilnehmer sowie 
    die Inhalte der ILIAS-Variablen zu den einzelnen Fragen ein. 
    Zusätzlich wird aus den Fragentiteln die FragenID, 
    also die Nummer der Frage extrahiert. 
    Weiterhin werden Vor- und Familienname extrahiert. 
    Die Annahme dabei ist, dass im Namen erst der Familienname kommt, 
    dann durch ", "(Komma und Spc) getrennt der Vorname.
    Die Ergebnisse sind in 1D bzw 2D Listen verfügbar.
"""

# Konstanten ###############################################################
Anz_Fragen = 14 # Anzahl Fragen pro Teilnehmer
Anz_Var = 15    # Maximale Anzahl von Frage-Variablen pro Frage
Anz_Res = 5     # Maximale Anzahl von Frage-Ergebnissen pro Frage. Fest = 1 in deiser Version
Max_Pkt = 58    # Maximale Anzahl von Punkten im Test
Schema_Note = ["5,0","4,0","3,7","3,3","3,0","2,7","2,3","2,0","1,7","1,3","1,0"]
Schema_Proz = [0,     25,   27,   29,   31,   33,   35,   37,   39,   41,   43]
              # Mindestprozentzahl an Punkten für die korrespondierende Note
Filename_Export = "TestergebnisseGE2.xlsx"
Filename_Export_public = 'ILIAS_TestergebnisseGE2.xlsx'
Filename_Export_PSSO = 'ILIAS_FragenpoolGE2.xlsx'
IRT_Frame_Name = "irt_frame_ge2.xlsx"

Name_Marker = "Ergebnisse von Testdurchlauf 1 für "
Formelfrage_Marker   = "Formelfrage"
Freitextfrage_Marker = "Freitext eingeben"
Freitextergebnis_Marker = "Ergebnis"
Vers = "1.11, 18.9.2020"
############################################################################

import pandas as pd
from math import * 

def Notenberechnung (Pkt, Max_Pkt, Schema_Proz, Schema_Note):
    """ E.Waffenschmidt, 4.9.2020
    Berechnet für eine erzielte Punktzahl einer Prüfung 
    anhand eines Notenschemas die Note
    """
    Note = "n.v." # Default Note, wenn nichts gefunden wird. Eigentlich unmöglich.
    for i in range(len(Schema_Note)):
       if (Pkt/Max_Pkt*100)>=Schema_Proz[i]:
           Note = Schema_Note[i]
    return Note

def Init_2D_None (m,n):
    """ E. Waffenschmidt, 3.9.2020
    initialisiert ein 2D-Array mit "None" als Inhalt
    Zugriff mit x = a[m][n]
    """
    a = []
    for i in range(m):
        a.append([None] * n)
    return a

def Init_2D_NoStr (m,n):
    """ E. Waffenschmidt, 3.9.2020
    initialisiert ein 2D-Array mit Leerstring als Inhalt
    Zugriff mit x = a[m][n]
    """
    a = []
    for i in range(m):
        a.append([""] * n)
    return a

def Init_3D_None (m,n,o):
    """ E. Waffenschmidt, 3.9.2020
    initialisiert ein 3D-Array mit "None" als Inhalt
    Zugriff mit x = a[m][n][o]
    """
    a = []
    for k in range(m):
        a.append(Init_2D_None(n,o))
    return a

def Init_3D_NoStr (m,n,o):
    """ E. Waffenschmidt, 3.9.2020
    initialisiert ein 3D-Array mit Leerstring als Inhalt
    Zugriff mit x = a[m][n][o]
    """
    a = []
    for k in range(m):
        a.append(Init_2D_NoStr(n,o))
    return a


def Get_Frage_ID (Titel):
    """ E. Waffenschmidt, 3.9.2020
    Extrahiert die Fragen-ID (Die "Nummer" der Frage) aus dem gesamten Titel. 
    Konkret sind das alle Zeichen bis zum ersten Leerzeichen.
    """
    return Titel[0:Titel.find(" ")] #Übernimmt den Text in Titel bis zum ersten Leerzeichen

def Finde_Fragenindex (FragenID, FragenID_Pool):
    """ E. Waffenschmidt, 3.9.2020
    Ermittelt den Index einer Frage in einem Fragebpool anhand der Fragen-ID
    Wenn die gesuchet ID nicht im Pool ist, wird "None" zurück geliefert.
    """
    Index = None
    Fragen_Anz = len(FragenID_Pool)
    for ID_Nr in range(0,Fragen_Anz):
        if FragenID == FragenID_Pool[ID_Nr]:
            Index = ID_Nr
    return Index

def Finde_Element (x, Liste):
    """ E. Waffenschmidt, 18.9.2020
    Sucht den Index von Element x in einer Liste.
    Wenn das Element nicht vorhanden ist, wird "None" zurückgeliefert
    Es sucht die ganze Liste durch. 
    Der Index zeigt auf den letzten gefunden Eintrag    
    """
    Index = None
    for i in range(0,len(Liste)):
        if x == Liste[i]:
            Index = i
    return Index

def eval_ILIAS (Gleichung_Ilias,v):
    """E. Waffenschmidt, 3.9.2020
       Evaluiert (d.h. nutzt die Gleichung zur Berechnung) 
       eine Formel im ILIAS-Format 
       Die Variablenwerte werden in Form einer Liste in der Variablen v übergeben.
       Die Anzahl der Variablen in der Liste ist beliebig 
       und ergibt sich aus der Länge der List-Variablen v.   
       ACHTUNG: Die erste Variable $v1 wird zu v[0].
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
       ,            .        manchmal taucht ein "," als Dezimaltrennzeichen auf
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
    Anz_Var = len(v) # Anzahl der möglichen Variable $v1, $v2 usw. in der ILIAS-Formel
# Gleichung für Python-Format anpassen:    
    Gleichung_Py = Gleichung_Ilias.lower()
    # Variablenbezeichnung umändern:
    for i in range(1,Anz_Var+1):
       Var_Ilias = "$v"+str(i)
       Var_Py = "v["+str(i-1)+"]"
       Gleichung_Py = Gleichung_Py.replace(Var_Ilias, Var_Py)
    # Mathematische Funktionen anpassen:
    Gleichung_Py = Gleichung_Py.replace(",", ".")
    Gleichung_Py = Gleichung_Py.replace("^", "**")
    Gleichung_Py = Gleichung_Py.replace("arcsin", "asin")
    Gleichung_Py = Gleichung_Py.replace("arcsinh", "asinh")
    Gleichung_Py = Gleichung_Py.replace("arccos", "acos")
    Gleichung_Py = Gleichung_Py.replace("arccosh", "acosh")
    Gleichung_Py = Gleichung_Py.replace("arctan", "atan")
    Gleichung_Py = Gleichung_Py.replace("arctanh", "atanh")
    Gleichung_Py = Gleichung_Py.replace("ln", "log")
    Gleichung_Py = Gleichung_Py.replace("log", "log10")

# Gleichung checken und berechnen
    try: # Testet, ob eine Fehler auftritt
       eval (Gleichung_Py)
    except: # Wenn ein Fehler auftritt, Fehlermeldung 
       Result = None
#       print ("!!! Gleichung ",Gleichung_Ilias," enthält einen Fehler:")
#       print ("    Python-Format:",Gleichung_Py)
#       print ("    Variablen:",v)
    else: # sonst is alles OK, und die Gleichung wird berechnet.
       Result = eval (Gleichung_Py)
    return Result

########################################################################
## HAUPTPROGRAMM ##############################################################
########################################################################

print ("Tool zur externen Bewertung von ILIAS Formelfragen-Tests")
print ("Version", Vers)
print ("(c) by Eberhard Waffenschmidt, TH-Köln")

# weitere Konstanten
Dummytext = "xyz"

# Daten aus EXCEL-File einlesen:
# Vor dem Filename in '' muss ein "r" gesetzt werden.
# Das Blatt im EXCEL-File wird nochmal mit sheet_name benannt
# header=None wird gesetzt, wenn keine Spaltenüberschriften existieren.
# skiprows=5 überspringt die ersten 5 Zeilen 

# Testergebnisse einlesen:
#df1 = pd.read_excel (r'C:\Users\Ebi\Documents\python\ILIAS Prüfungs-Auswertung\ETAT-Probe-Klausur_results kurz.xlsx', sheet_name='Auswertung für alle Benutzer')
#df1 = pd.read_excel (r'C:\Users\Ebi\Documents\python\ILIAS Prüfungs-Auswertung\ETAT-Probe-Klausur_results mit Titelzeile.xlsx', sheet_name='Auswertung für alle Benutzer', header=None)
print ("EXCEL-Daten werden eingelesen...")
df1 = pd.read_excel (r'ILIAS_Testergebnisse.xlsx', sheet_name='Auswertung für alle Benutzer', header=None)
print ("ILIAS-Ergebnisse OK.")
# Fragenpool aus Excel-Tabelle einlesen
df2 = pd.read_excel (r'ILIAS_Fragenpool.xlsx', sheet_name='Tabelle1', skiprows=6)
#df2 = pd.read_excel (r'C:\Users\Ebi\Documents\python\ILIAS Prüfungs-Auswertung\EGT-ILIAS-Klausuraufgaben  edEW-FK4_1.xlsx', sheet_name='Tabelle1', skiprows=5)
print ("EXCEL-Fragenpool OK.")
# Bonuspunkte aus Excel-Tabelle einlesen
df3 = pd.read_excel (r'Bonuspunkte.xlsx', sheet_name='Tabelle1')
print ("Bonuspunkte OK.")
# PSSO-Teilnehmerliste einlesen
df4 = pd.read_excel (r'PSSO_Teilnehmer.xls', sheet_name='First Sheet', skiprows=3)
print ("PSSO Teilnehmerliste OK.")

########################################################################
### Testergebnisse verarbeiten: ########################################
########################################################################
# Daten in 2D-Array umwandeln, damit der Zugriff einfacher zu indexieren ist
I_D = df1.values #I_D steht für ILIAS-Daten
Anz_Zeilen = len(I_D)

# Anzahl Teilnehmer ermitteln
Anz_Teilnehmer = 0
for Zeile in range(Anz_Zeilen):
   txt = I_D[Zeile,0]
   if (txt.__class__ == Dummytext.__class__):
      if txt.startswith(Name_Marker):
         Anz_Teilnehmer = Anz_Teilnehmer+1

# Variable initialisieren
Nr_Teilnehmer = [None] * Anz_Teilnehmer # Fortlaufende Nummer 
Namen = [""] * Anz_Teilnehmer
Vornamen = [""] * Anz_Teilnehmer
Familiennamen = [""] * Anz_Teilnehmer
MatNr = [""] * Anz_Teilnehmer
Noten = [""] * Anz_Teilnehmer
Fragentitel = Init_2D_NoStr(Anz_Teilnehmer,Anz_Fragen) # Fragentitel = [[""]*Anz_Fragen]*Anz_Teilnehmer fuktioniert nicht!
FragenID = Init_2D_NoStr(Anz_Teilnehmer,Anz_Fragen)
Fragentypen = Init_2D_NoStr(Anz_Teilnehmer,Anz_Fragen) # Fragentypen "Formelfrage" und "Freitextfrage" werden unterstützt
Fragen_Formel = Init_2D_NoStr(Anz_Teilnehmer,Anz_Fragen)
Fragen_Tol = Init_2D_None(Anz_Teilnehmer,Anz_Fragen)
Var = Init_3D_None(Anz_Teilnehmer,Anz_Fragen,Anz_Var) # Zugriff mit: V = Var[Teilnehmer][Frage][Variable]
Res = Init_3D_None(Anz_Teilnehmer,Anz_Fragen,Anz_Res)
Res_Ref = Init_3D_None(Anz_Teilnehmer,Anz_Fragen,Anz_Res) # Richtiges Ergebnis als Referenz
Pkt = Init_3D_None(Anz_Teilnehmer,Anz_Fragen,Anz_Res)     # Vergebene Punkte für die Aufgabe
Bonuspunkte = [0] * Anz_Teilnehmer
Ges_Pkt = [None] * Anz_Teilnehmer
Fragentyp = "unbekannt"

# Daten der Teilnehmer analysieren
print("Testergebnisse werden analysiert...")
Teilnehmer = 0
Frage_Nr = 0
for Zeile in range(Anz_Zeilen):
   txt = I_D[Zeile,0]
   if (txt.__class__ == Dummytext.__class__):
      if txt.startswith(Name_Marker):
         Teilnehmer = Teilnehmer+1
         Nr_Teilnehmer[Teilnehmer-1] = Teilnehmer
         Name = txt.replace(Name_Marker,"") # Der Text vor dem Namen wird entfernt
         Namen[Teilnehmer-1] = Name
         Familiennamen[Teilnehmer-1] = Name[0:Name.find(",")] #Übernimmt den Text in Titel bis zum ersten Komma.
         Vornamen[Teilnehmer-1] = Name[Name.find(",")+2:]     #Übernimmt den Text in Titel 2 Stellen nach dem ersten Komma.
         # print ("Teilnehmer Nr.: ",Teilnehmer, " : ",Name," Vorname: ",Vornamen[Teilnehmer-1], " Familienname: ",Familiennamen[Teilnehmer-1] )
         Frage_Nr = 0
         Fragentyp = "unbekannt"

      if Fragentyp == "Formelfrage":
         for Var_Nr in range(1,Anz_Var+1):
             Var_Marker = "$v"+str(Var_Nr)
             if txt.startswith(Var_Marker):
                x = I_D[Zeile,1]
                Var[Teilnehmer-1][Frage_Nr-1][Var_Nr-1] = x
    #            print ("Var[",Var_Nr,"] = ",x)
         for Res_Nr in range(1,Anz_Res+1):
             Res_Marker = "$r"+str(Res_Nr)
             if txt.startswith(Res_Marker):
                y = I_D[Zeile,1]
                Res[Teilnehmer-1][Frage_Nr-1][Res_Nr-1] = y
    #            print ("Res[",Res_Nr,"] = ",y," - ",Res[Teilnehmer-1][Frage_Nr-1], "Vorh.R=",Res[Teilnehmer-1][Frage_Nr-2],)
      if txt.startswith(Formelfrage_Marker):
         Frage_Nr = Frage_Nr + 1
         Fragentyp = "Formelfrage"
         Fragentypen[Teilnehmer-1][Frage_Nr-1] = Fragentyp
         Titel = I_D[Zeile,1]
         Fragentitel[Teilnehmer-1][Frage_Nr-1] = Titel
         FragenID[Teilnehmer-1][Frage_Nr-1] = Get_Frage_ID (Titel)
#         print ("Teilnehmer Nr.: ",Teilnehmer, " Fragenummer: ",Frage_Nr, " ID: ",FragenID[Teilnehmer-1][Frage_Nr-1])

      if Fragentyp == "Freitextfrage": #Aktueller Fragentyp ist Freitextfrage
         if txt.startswith(Freitextergebnis_Marker): #in der LinkenSpalte steh der Marker für Freittextergebnis (default "Ergebnis")
            Res_Nr = Res_Nr+1          # vielleicht gibt es mehr als ein Ergebnis. Dann wird hier hochgezählt
            y = I_D[Zeile,1]           # Der Inhalt in der rechten (zweiten) Spalte wird jetzt als Ergebnis übernommen
#            print("Teiln.",Teilnehmer,", Frage",Frage_Nr, "Freitext:",y)
            Res[Teilnehmer-1][Frage_Nr-1][Res_Nr-1] = y
      if txt.startswith(Freitextfrage_Marker):
         Frage_Nr = Frage_Nr + 1
         Res_Nr = 0
         Fragentyp = "Freitextfrage"
         Fragentypen[Teilnehmer-1][Frage_Nr-1] = Fragentyp
         Titel = I_D[Zeile,1]
         Fragentitel[Teilnehmer-1][Frage_Nr-1] = Titel
         FragenID[Teilnehmer-1][Frage_Nr-1] = Get_Frage_ID (Titel)
      
########################################################################
### Fragenpool verarbeiten #############################################
########################################################################
# relevante Spalten in 1D-Arrays (Lists) kopieren
Questiontypes_Pool = df2['Question Type'] # "Formelfrage" oder "Freitextfrage"
Titels_Pool = df2['Question Title']
Gleichungen1_Pool = df2['res1_formula'] # Nur Formel 1 wird ausgewertet
Toleranzen1_Pool = df2['res1 tol']
Punkte1_Pool = df2['res1 pts']
Anz_Fragen_Pool = len(Titels_Pool)

FragenID_Pool = [""] * Anz_Fragen_Pool
# Fragen IDs aus den Titeln extrahieren:
for i in range(0,Anz_Fragen_Pool):
    FragenID_Pool[i] = Get_Frage_ID(Titels_Pool[i])
    
########################################################################
### Testergebnisse und Fragenpool verknüpfen ###########################
########################################################################

# Formeln zu den Ergebnissen der Teilnehmer zuordnen, 
# Richtiges Ergebnis berechnen 
# Mit den Ergebnis des Studenten vergleichen 
# und Punkte vergeben    
print("Testergebnisse werden mit Fragenpool verknüpft...")

for Teilnehmer in range(0,Anz_Teilnehmer):
    Ges_Punkte = 0
    for Frage in range(0,Anz_Fragen):
        Frage_OK = True # Frage erstmal per Default auf OK setzen. Wird genutzt um buggy Aufgaben "per Hand" rauszufiltern.
        Fragen_Punkte_Student = 0
        # Index der Frage im Fragenpool finden
        i = Finde_Fragenindex (FragenID[Teilnehmer][Frage], FragenID_Pool)
        if i==None: #Falls Frage nicht im Pool gefunden wird
            R = -999999
            Formel = "Formel Nicht gefunden"
            Toleranz = None
            print ("!!! Teiln.",Teilnehmer,", Frage",Frage,",",FragenID[Teilnehmer][Frage]," existiert nicht im Pool!")
        else: 
            if Questiontypes_Pool[i] == "Formelfrage": # Frage ist eine Formelfrage
                # Passende Formel usw. aus dem Fragenpool auslesen
                # Derzeit wird nur Ergebnis 1 ausgewertet
                Formel = Gleichungen1_Pool[i]
                Toleranz = Toleranzen1_Pool[i]
                Punkte = Punkte1_Pool[i]
                
                # Formel mit den Variablen des Studenten anwenden:
                v = Var[Teilnehmer][Frage]
                if v[0] == None: # Wenn der Student die Aufgabe gar nicht angeschaut hat, sind alle Variablen None, insbesondere die erste.
                    R = None    
                else: # Der Student hat wenigstens die Frage angeschaut und Werte bekommen
                    
############################## Spezial ETAT 2020-10: ######################################
                    # Manuelles Rausfiltern von buggy Aufgaben
                    # Prüfen, ob die Frage eine von den fehlerhaften Fragen ist
#                    print ("!!! Teiln.",Teilnehmer,", Frage",Frage,",",FragenID[Teilnehmer][Frage],"wird geprüft")
                    if FragenID[Teilnehmer][Frage] == '07.1.04':
                        # Frage enthält Serienschaltung. 
                        # Gesamtspannung v2 in mV ist manchmal kleiner als Einzelspannung v3in V.
                        # Das ist physikalisch nicht möglich, 
                        # Die Lösung mit der angegebenen Formel ist dann Quatsch.
                        # Daher wird falls v2 < v3 ist, ein Fehler erkannt
                        # und immer ein Gnadenpunkt vergeben. 
                        # und erst gar nicht die Formel berechnet. 
                        # Stattdessen ist das Ergebnis 0.
                        if v[1]/1000 < v[2]:
                            Frage_OK = False
                            Fragen_Punkte_Student = Punkte
                            Ges_Punkte = Ges_Punkte + Punkte
                            R = 0
                            print("Aufgabe 07.1.04: U_0 =",v[1],"mV < U_2 =",v[2],"V -> Gnadenpunkt vergeben")
##########################################################################################
                    if Frage_OK:
                        R = eval_ILIAS (Formel,v)
                        # print ("Formel =",Formel,"=",R,"Toleranz:",Toleranz,"%")
                        if R != None: #Wenn die Formel ein sinnvolles Ergebnis liefert:
                           # SONDERBEHANDLUNG von Frage 22.1.03:
                           if FragenID[Teilnehmer][Frage] == '22.1.03':
                               # Bei der Lösung kann ein Winkel >360 rauskommen, da die Ranges schlecht 
                               # gewählt sind. Es werden nun Studi-Lösung und Python-Lösung jeweils
                               # in den Bereich 0..360° gebracht und dann erst verglichen.
                               # Weiter unten gibt es nochmal eine extra Abfrage 
                               # nur für diese Frage
                               R = remainder (R, 360)

                           # Maximale und minimale Grenze mit Toleranz bestimmen.
                           # ACHTUNG: Bei negativem Vorzeichen drehen sich min und max rum, 
                           # das gibt dann Ärger beim nachfolgenden Vergleich 
                           # Daher hier die Verwendung von "min" und "max"
                           R_min = min (R*(1+Toleranz/100),R*(1-Toleranz/100))
                           R_max = max (R*(1+Toleranz/100),R*(1-Toleranz/100))
                           # Ergebnis des Studierenden:
                           R_Student = Res[Teilnehmer][Frage][0] # Es wird in dieser version nur ein Ergebnis, das erste, ausgewertet 
                           # Ist das Ergebnis vorhanden und innerhalb der Toleranz?
                           # Dann gibt's die Punkte für die Aufgabe, sonst 0 Pkt.
                           if R_Student != None: #Wenn der Student keine Lösung angegeben hat ist R_Student = None
                               # Dann kann es noch sein, dass die Lösung als Bruch, z.B. 1/300 angegebn ist.
                               # Dann muss der Bruch mit eval ausgerechnet werden
                               if type(R_Student)==str:
                                   R_Student = eval(R_Student)
############################## Spezial ETAT 2020-10: ######################################
                               if FragenID[Teilnehmer][Frage] == '22.1.03':
                                   # Bei der Lösung kann ein Winkel >360 rauskommen, da die Ranges schlecht 
                                   # gewählt sind. Es werden nun Studi-Lösung und Python-Lösung jeweils
                                   # in den Bereich 0..360° gebracht und dann erst verglichen.
                                   # Weiter oben gibt es schon mal eine extra Abfrage 
                                   # nur für diese Frage, und unten eine Abfrage zum Print.
                                   R_Student = remainder (R_Student, 360)
##########################################################################################
                               if (R_Student >= R_min) and (R_Student <= R_max):
                                   Fragen_Punkte_Student = Punkte
                                   Ges_Punkte = Ges_Punkte + Punkte
############################## Spezial ETAT 2020-10: ######################################
                                   if FragenID[Teilnehmer][Frage] == '22.1.03':
                                       # Bei der Lösung kann ein Winkel >360 rauskommen, da die Ranges schlecht 
                                       # gewählt sind. 
                                       print("Aufgabe 22.1.03: Punkt vergeben")
##########################################################################################
                        else: # Formel war fehlerhaft. Daher: Im Zweifel für den Angeklagten, volle Punktzahl
                           Fragen_Punkte_Student = Punkte
                           Ges_Punkte = Ges_Punkte + Punkte
                           
                # Jetzt noch die Ergebnisse in den Listen abspeichern:
                Fragen_Formel[Teilnehmer][Frage] = Formel
                Fragen_Tol[Teilnehmer][Frage] = Toleranz
                Res_Ref[Teilnehmer][Frage][0] = R
                Pkt[Teilnehmer][Frage][0] = Fragen_Punkte_Student
 
            if Questiontypes_Pool[i] == "Freitextfrage": # aktuelle Frage ist eine Freitextfrage
                # Freitext-Schlüsselwort aus dem Fragenpool auslesen
                # Derzeit wird nur Ergebnis 1 ausgewertet
                Schluesselwort = Gleichungen1_Pool[i]
                Punkte = Punkte1_Pool[i]
                R_Student = Res[Teilnehmer][Frage][0] # Es wird in dieser version nur ein Ergebnis, das erste, ausgewertet 
                # Ist das Ergebnis vorhanden? 
                if R_Student != None: #Wenn der Student keine Lösung angegeben hat ist R_Student = None
                   if R_Student.__class__ == Dummytext.__class__: # und das Ergebnis vom Typ Text ist
                       # Ist das Schlüsselwort in der Antwort des Studi enthalten?
                       if Schluesselwort in R_Student:   
                          Fragen_Punkte_Student = Punkte # Dann gibt's die Punkte für die Frage
                          Ges_Punkte = Ges_Punkte + Punkte
                          Fragen_Formel[Teilnehmer][Frage] = Schluesselwort
                          Pkt[Teilnehmer][Frage][0] = Fragen_Punkte_Student
            
    # Gesamtpunkte und Note für Teilnehmer verbuchen
    Ges_Pkt[Teilnehmer] = Ges_Punkte

########################################################################
### Bonuspunkte hinzufügen #############################################
########################################################################

print("Bonuspunkte werden hinzugefügt...")
Bonus_Namen = df3['Name']
Bonus_Punkte_Liste = df3['Bonuspunkte']

for Teilnehmer in range(0,Anz_Teilnehmer):
    Bonus_Index = Finde_Element (Namen[Teilnehmer], Bonus_Namen) # finde Namen in der Bonuspunkteliste
    if Bonus_Index != None: #Wenn der Name gefunden wurde, werden die Punkte vergeben.
        Bonuspunkte[Teilnehmer] = Bonus_Punkte_Liste[Bonus_Index]
        Ges_Pkt[Teilnehmer] = Ges_Pkt[Teilnehmer] + Bonuspunkte[Teilnehmer]

########################################################################
### Noten berechnen ####################################################
########################################################################

print("Noten werden berechnet...")
for Teilnehmer in range(0,Anz_Teilnehmer):
    Noten[Teilnehmer] = Notenberechnung (Ges_Pkt[Teilnehmer], Max_Pkt, Schema_Proz, Schema_Note)
    print ("Nr.",Nr_Teilnehmer[Teilnehmer],Namen[Teilnehmer],", Ges.Pkt =",Ges_Pkt[Teilnehmer],", Note =",Noten[Teilnehmer])

print ("Anzahl Teilnehmer     = ", Anz_Teilnehmer)
print ("Anzahl Fragen im Pool = ", Anz_Fragen_Pool)

########################################################################
### Matrikelnummern ermitteln ##########################################
########################################################################
print("Matrikelnummern werden ermittelt...")

PSSO_Vornamen = df4['vorname']
PSSO_Familiennamen = df4['nachname']
PSSO_MatNr = df4['mtknr']        # MatrNr aus dem PSSO, teilweise Text, teilweise Zahl
Anz_PSSO = len(PSSO_Vornamen)
PSSO_Namen = [""] * Anz_PSSO
PSSO_MatNr_Txt = [""] * Anz_PSSO # wird die MatNr aus dem PSSO in Text-Form enthalten
Dummyzahl = 1

# Manchmal ist MatNr eine Zahl und kein Text.
# Einheitlich in Text umwandeln
for i in range(0,Anz_PSSO):
    MN = PSSO_MatNr[i]
    if MN.__class__ == Dummyzahl.__class__:
        PSSO_MatNr_Txt[i] = str(MN)
    else:
        PSSO_MatNr_Txt[i] = MN

PSSO_MatNr = PSSO_MatNr_Txt
df4['mtknr'] = PSSO_MatNr_Txt

# in der PSSO-Liste sind die Sortnamen bei langen Namen verkrüppelt.
# Daher volle Namen aus Vor- und Familienname zum Vergleich zusammensetzen
# Getrennt mit ", " wie in der ILIAS-Liste
for i in range(0,len(PSSO_Vornamen)):
    # Offensichtlich pfuschen sich manchmal Leerzeilen mit NaN dazwischen. 
    # Die sind keine STRG-Klasse und werden übersprungen
    if (PSSO_Vornamen[i].__class__== Dummytext.__class__) and (PSSO_Familiennamen[i].__class__ == Dummytext.__class__):
        PSSO_Namen[i] = PSSO_Familiennamen[i] + ", " + PSSO_Vornamen[i]

for Teilnehmer in range(0,Anz_Teilnehmer):
    MatNr_Index = Finde_Element (Namen[Teilnehmer], PSSO_Namen) # finde Namen in der PSSO_Liste
    if MatNr_Index != None: #Wenn der Name gefunden wurde, wird die Matrikenummer gespeichert
        MN = PSSO_MatNr[MatNr_Index]
        # Manchmal ist MatNr eine Zahl und kein Text.
        # Einheitlich in Text umwandeln
        if MN.__class__ == Dummyzahl.__class__:
            MatNr[Teilnehmer] = str(MN)
        else:
            MatNr[Teilnehmer] = MN
    else: # MatNr wurde nicht gefunden, Meldung ausgeben
        print ("!!! >",Namen[Teilnehmer],"< ist nicht in der PSSO-Liste zu finden!")

########################################################################
### Noten der PSSO-Liste zuordnen ######################################
########################################################################
print ("Noten werden der PSSO-Liste zugeordnet...")
PSSO_Bewertung = [""] * Anz_PSSO # Alle Noten per Default erstmal leer setzen
for i in range(0,Anz_PSSO):
    MatNr_Index = Finde_Element (PSSO_MatNr_Txt[i], MatNr) # Finde Mat.Nr in der Ergebnis-Liste
    if MatNr_Index != None: #Wenn der MatNr gefunden wurde, wird die Note gespeichert
        PSSO_Bewertung[i] = Noten[MatNr_Index]
    else: # Wenn nicht, hat Person nicht teilgenommen, Eintrag PNE (Prüfung nicht erschienen)
        PSSO_Bewertung[i] = "PNE"
        print (PSSO_Vornamen[i]," ",PSSO_Familiennamen[i],", MatNr.",PSSO_MatNr_Txt[i],", ist nicht erschienen.")
# Spalte mit Noten dem Dataframe zuordnen:
df4['bewertung'] = PSSO_Bewertung

########################################################################
### Daten in EXCEL-Sheet exportieren  ##################################
########################################################################
# Dazu passenden Pandas-Dataframe zusammenbauen:
# Zeilentitel generieren:
#  Nr Name Vorname Familienname MatNr Note GesPkt A1Pkt ... A40Pkt ...
#  ... A1_ID A01_Formel A01_Tol A1_ResRef A1_Res  A1_v1 A1_v2...A1_v10 ...
#  ...
#  ... A40_ID A01_Formel A40_Tol A40_ResRef A40_Res  A40_v1 A40_v2...A40_v10
    
print ('Daten werden nach EXCEL exportiert...')

DF_ex = pd.DataFrame() # leeren DataFrame zum Export erzeugen
DF_public = pd.DataFrame() # leeren DataFrame zum Export der öffentlichen (public) Daten erzeugen
# Übersichtsdaten
DF_ex['Nr'] = Nr_Teilnehmer  # erzeugt eine neue Spalte mit dem Titel 'Nr' und Daten in Nr_Teilnehmer
DF_ex['Name'] = Namen  # erzeugt eine neue Spalte mit dem Titel 'Name' und den Daten in Namen
DF_ex['Vorname'] = Vornamen
DF_ex['Familienname'] = Familiennamen 
DF_ex['MatNr'] = MatNr
DF_ex['Note'] = Noten
DF_ex['GesPkt'] = Ges_Pkt
DF_ex['BonusPkt'] = Bonuspunkte

DF_public['MatNr'] = MatNr
DF_public['Note'] = Noten
DF_public['GesPkt'] = Ges_Pkt
DF_public['BonusPkt'] = Bonuspunkte

# Gesamtpunkte bei den einzelnen Fragen
for Frage in range (Anz_Fragen):
    Spaltentitel = "A"+str(Frage+1)+"_Pkt"
    x = [None]*Anz_Teilnehmer # Spaltendaten initialisieren
    for Teilnehmer in range(Anz_Teilnehmer):
        x[Teilnehmer] = Pkt[Teilnehmer][Frage][0]
    DF_ex [Spaltentitel] = x
    DF_public [Spaltentitel] = x

DF_ex[''] = [""]*Anz_Teilnehmer # Leerspalte an dieser Stelle einfügen
DF_public[''] = [""]*Anz_Teilnehmer # Leerspalte an dieser Stelle einfügen

# Details zu den einzelnen Fragen 
for Frage in range (Anz_Fragen):
    Spaltentitel = "A"+str(Frage+1)+"_ID"
    x = [""]*Anz_Teilnehmer # Spaltendaten initialisieren
    for Teilnehmer in range(Anz_Teilnehmer):
        x[Teilnehmer] = FragenID[Teilnehmer][Frage]
    DF_ex [Spaltentitel] = x

    Spaltentitel = "A"+str(Frage+1)+"_Formel"
    x = [""]*Anz_Teilnehmer # Spaltendaten initialisieren
    for Teilnehmer in range(Anz_Teilnehmer):
        x[Teilnehmer] = Fragen_Formel[Teilnehmer][Frage]
    DF_ex [Spaltentitel] = x

    Spaltentitel = "A"+str(Frage+1)+"_Tol"
    x = [""]*Anz_Teilnehmer # Spaltendaten initialisieren
    for Teilnehmer in range(Anz_Teilnehmer):
        x[Teilnehmer] = Fragen_Tol[Teilnehmer][Frage]
    DF_ex [Spaltentitel] = x
    
    Spaltentitel = "A"+str(Frage+1)+"_Res_Ref"
    x = [""]*Anz_Teilnehmer # Spaltendaten initialisieren
    for Teilnehmer in range(Anz_Teilnehmer):
        x[Teilnehmer] = Res_Ref[Teilnehmer][Frage][0]
    DF_ex [Spaltentitel] = x
    DF_public [Spaltentitel] = x
    
    Spaltentitel = "A"+str(Frage+1)+"_Res"
    x = [""]*Anz_Teilnehmer # Spaltendaten initialisieren
    for Teilnehmer in range(Anz_Teilnehmer):
        x[Teilnehmer] = Res[Teilnehmer][Frage][0]
    DF_ex [Spaltentitel] = x
    DF_public [Spaltentitel] = x
    
    # Variablen der einzelnen Fragen
    for Variable in range(Anz_Var):
        Spaltentitel = "A"+str(Frage+1)+"_v"+str(Variable+1)
        x = [""]*Anz_Teilnehmer # Spaltendaten initialisieren
        for Teilnehmer in range(Anz_Teilnehmer):
            x[Teilnehmer] = Var[Teilnehmer][Frage][Variable]
        DF_ex [Spaltentitel] = x    
        DF_public [Spaltentitel] = x    
    
# Datenframe in EXCEL-File schreiben   
DF_ex.to_excel(Filename_Export, index=False) # Index = False sorgt dafür, dass die erste Spalte nicht den Zeilenidex von 0..Ende enthält
DF_public.to_excel(Filename_Export_public, index=False) # Index = False sorgt dafür, dass die erste Spalte nicht den Zeilenidex von 0..Ende enthält
df4.to_excel(Filename_Export_PSSO, index=False) # Index = False sorgt dafür, dass die erste Spalte nicht den Zeilenidex von 0..Ende enthält 
print ('Fertig!')

pkt_per_frage = {}
for uid, fids, pkte in zip(range(Anz_Teilnehmer), FragenID, Pkt):
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

pd.DataFrame(irtf).T.fillna(0).to_excel(IRT_Frame_Name)