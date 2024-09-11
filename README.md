# Luftlinientool2022
Erstellen von Luftlinienverbindungen für gegebene Bezirke anhand der RIN.

## Anwendung
### Anwendung via Dialog
Zwei Möglichkeiten
* Visumversion ist noch geschlossen:
  
  GUI extern starten (*llt_GUI.py* ausführen)
* Visumversion ist bereits geöffnet:
  * Das Ausführen der GUI (*llt_GUI.py*) in das Skriptmenü integrieren
  * Skript via Skriptmenü starten 
    
    Hinweis: Wenn die GUI mehrmals gestarten & beendet wird, erscheint eine Fehlermeldung. Diese kann ignoriert werden, die Funktionalität ist trotzdem gegeben. 

### Aufruf via Code
Das Luftlinientool kann auch ohne GUI angewendet werden. Dazu muss als Codeausführung eine Instanz der Klasse LuftlinienCalculator erstellt werden.
Danach kann auf die Methoden der Instanz (Import, Berechnung, Export) zugegriffen werden
Ein Beispiel ist unter *Bsp_Aufruf_ohne_GUI.py* zu sehen.

### auszuführende Schritte
1. Parameter setzen (welche VFS, Attributswerte etc.)

2. Luftlinienkalkulatorobjekt initialisieren
   
    Code: Aufruf Konstruktor mit Parameterübergabe
   
    GUI: "Daten einlesen" in Toolbar ausführen
   
3. Luftlinienverbindungen berechnen/erzeugen

    Code: Aufruf calculate_main Methode
   
    GUI: "Berechnung Luftlinien-Netz" in Toolbar ausführen
   
4. Ergebnisse in gewünschter Form nach Visum exportieren

    Code: Aufruf der export_matrix/export_net Methode
   
    GUI: Die entsprechenden Buttons (Mtx/Net) in der Spalte "anlegen in Visum als" verwenden. Alternativ überträgt der Button "Import nach Visum alle VFS Strecken + Mtx" die kombinierten Ergebnisse aller VFS.
  
Anmerkungen:

* Bei der GUI werden aktuelle Berchnungen zurückgesetzt, wenn die Auswahl eines der Bezirksattribtue geändert wird. Dabei wird eine neue Instanz des LLT Kalkulators erstellt.
* Werden Bezirkswerte in Visum geändert, werden diese nicht automatisch im LLT Kalkulator geändert. Deshalb muss der Verfahrensablauf ab Schritt 2 wieder ausgeführt werden.
* Die initialen Parameterwerte können in der GUI über das Tool "Defaultwerte" wieder aufgerufen werden
* "Ergebnisse initialisieren" ermöglicht das Löschen bereits vorhandener Ergebnisse
* Schritt 3 verwendet die aktuell in der GUI eingegebenen Parameter. Vor der Rechnung mit neuen Parametern empfiehlt sich das Löschen der vorhandenen Ergebnisse ("Ergebnisse initialisieren"). 

## Vorraussetzung
Netz mit kategorisierten Bezirken:
* Attribut für die Zentralität (Bezirke): je kleiner die Zahl, desto größer ist die Zentralität des Bezirks

    Bsp.: Angabe der Zentralität über die Typnummer
    * 0 ... Metropolregion
    * 1 ... Oberzentrum
    * 2 ... Mittelzentrum
    * 3 ... Grundzentrum
    * 4 ... Ort ohne zentrale Funktion
    * 5 ... Teilort
    
* (optional) aktiver Bezirksfilter
* (optional) Angabe eines Attributs, welcher Bezirk als Quelle verwendet werden soll) {0=Nein, 1=Ja}
* (optional) Angabe eines Attributs, welcher Bezirk als Ziel verwendet werden soll) {0=Nein, 1=Ja}
 
 
