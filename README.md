Eplan.Scripting.Prototypes
==========================

Zweck:
Erstellung eines neuen Projektes mit Voreinstellungen wie bei Basis- oder Vorlageprojekten, allerdings mit dynamisch befüllbaren 
Strukturkennzeichen-Beschreibungsfelder (PropID 1000, 1002, 1007, 1008, 1009_1 bis 1009_10). 

Bekannte Fehler:
  - nicht Compilierbar wegen fehlender Objektreferenz (Object reference not set to an instance of an object.)
    (ist schon spät)
  - momentan nur  für Anlagenkennzeichen vorgesehn
  - dutzende nicht behandelte Exceptions
  - nur ein Kennzeichen möglich (Iterationsschleife fehlt) 

Voraussetzungen: 
  - eine möglichts kleine EPJ Datei muss Vorhanden sein [epjsource]
  - das zu erstellende Projekt darf noch nicht existieren [absPrjName]
  - die gewünschten Projekteinstellungen müssen nach [settings] exportiert sein
  - die gewünschten Auswertungsvorlagen müssen nach [auswertungen] exportiert sein
  - im public void StartScript muss ein Objekt der Klasse NewProjectWithIdDescription.Strukturkennzeichen 
    instaziert, die Properties gesetzt und der Methode WriteEpjFile(string, Strukturkennzeichen) übergeben werden
  
Funktionsweise:
  - Klasse Strukturkennzeichen wird mit Anlagenkennzeichen als string instanziert
  - Beschreibungen können als Eigenschaften von Strukturkennzeichen in form von MultiLangStrings gesetzt werden
  - Aus der epjsource und der übergebenen Strukturkennzeichen wird eine neue epj geschrieben und als neues Projekt mit der import-Action     importiert
  - ProjectSettings werden in neu erzeugtes Projekt importiert (nötig wegen minimaler epj)
  - Auswertungen werden in neu erzeugtes Projekt importiert
  - Projekt wird in P8 geöffnet 
  - Action XPrjActionProjectCompleteMasterData wird ausgeführt (Projektstammdaten vervollständigen)
