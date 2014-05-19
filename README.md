Eplan.Scripting.Prototypes
==========================

Zweck:
Erstellung eines neuen Projektes mit Voreinstellungen wie bei Basis- oder Vorlageprojekten, allerdings mit dynamisch befüllbaren 
Strukturkennzeichen-Beschreibungsfelder (PropID 1000, 1002, 1007, 1008, 1009_1 bis 1009_10). 

Bekannte Fehler:
  - nicht Compilierbar wegen fehlender Objektreferenz (Object reference not set to an instance of an object.)
    

Voraussetzungen: 
  - eine möglichts kleine EPJ Datei muss Vorhanden sein [epjsource]
  - das zu erstellende Projekt darf noch nicht existieren [absPrjName]
  - die gewünschten Projekteinstellungen müssen nach [settings] exportiert sein
  - die gewünschten Auswertungsvorlagen müssen nach [auswertungen] exportiert sein
  - im public void StartScript muss ein Objekt der Klasse NewProjectWithIdDescription.Strukturkennzeichen 
    instaziert, die Properties gesetzt und der Methode WriteEpjFile(string, Strukturkennzeichen) übergeben werden
  
