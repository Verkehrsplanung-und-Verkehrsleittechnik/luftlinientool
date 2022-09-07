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
 
 
