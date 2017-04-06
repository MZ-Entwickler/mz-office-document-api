# MZ Office Document API
Bibliothek zum Ausfüllen von DOCX (Microsoft Word ab 2007) und ODT (Open-/LibrOffice) Dokumenten
die Platzhalter oder Tabellen (auch ineinander verschachtelte Tabellen) beinhalten.

Zur Erstellung der Dokumente wird ein einheitliches Daten-Modell (_com.mz.solutions.office.model_)
verwendet, das unabhängig des Dokumenten-Formates (docx/odt) verwendbar ist.

Für Microsoft Word liegt eine Erweiterung vor zur Nutzung von Custom XML Parts mit denen
Word-Dokumente XML gemapped werden können.

## Lizenz
Die Bibliothek ist unter den Bedingungen der Affero General Public License Version 3 (AGPL v3)
frei verfügbar und verwendbar.

## Abhängigkeiten
- _Runtime_: Es werden keine Abhängigkeiten zur Laufzeit benötigt/verwendet.
- _Compile_: JSR 305 (Annotations for Software Defect Detection). Die entsprechende JAR findet
  sich im _third-party_ Verzeichnis.
- _Test_: JUnit 4.0 und hamcrest-core 1.0. Beides liegt im _third-party_ Verzeichnis vor.

Es wird Java 1.8 benötigt.

# Beispiele
Im Verzeichnis 'examples' findest du min. ein Beispiel, dessen Code kommentiert ist. Ansonsten
müssten die JavaDoc-Kommentare (inkl. Package-Infos) doch recht ausführlich sein.

__Grundsätzlich__ sollte die JavaDoc-Doku  mehr als ausreichend und genau sein. Wenn du dir nicht nur
die Klassen anschaust, sondert auch die Package-Infos, dann ist dies doch relativ viel. :-)

# Einstieg - com.mz.solutions.office
Jeder Vorgang zum Befüllen von Textdokumenten (*.odt, *.docx, *.doct) bedarf eines
Dokumenten-Daten-Modelles das im Unterpackage 'model' zu finden ist. Nach dem Erstellen/
Zusammensetzen des Modelles oder laden aus externen Quellen, kann in weiteren Schritten mit 
diesem Package ein beliebiges Dokument befüllt werden.

## Dokumente öffnen
Alle Dokumente werden auf dem selben Weg geöffnet. Ist das zugrunde liegende Office-Format 
(ODT oder DOCX) bereits vor dem Öffnen bekannt, kann die konrekte Implementierung gewählt werden.

```java
  // für Libre/ Apache OpenOffice
  OfficeDocumentFactory docFactory = OfficeDocumentFactory.newOpenOfficeInstance();
 
  // für Microsoft Office (ab 2007)
  OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
```

Ist das Dokumentformat unbekannt, können die Implementierung selbst erkennen ob es sich um ihr
eigenes Format handelt.

```java
  Path anyDocument = Paths.get("unknownFormat.blob");
  Optional<Office> docFactory = OfficeDocumentFactory.newInstanceByDocumentType(anyDocument);
 
  if (docFactory.isPresent() == false) {
      throw new IllegalStateException("Unbekanntes Format");
  }
```

Die Wahl zur passenden Implementierung bedarf dabei keiner korrekten Angabe der 
Dateinamenserweiterung. Die OfficeDocumentFactory und deren konkrete Implementierung erfüllt 
immer das Dokumenten-Daten-Modell vollständig.

## Konfiguration der Factory
Mögliche Konfigurationseigenschaften die für alle Implementierungen gültig sind finden sich in 
OfficeProperty; Implementierungs-spezifische in eigenen Klassen (z.B. MicrosoftProperty). Die 
Factory ist nicht thread-safe und eine Änderung der Konfiguration wikt sich auf alle bereits 
geöffneten, zukünftig zu öffnenden und derzeit in Bearbeitung befindlichen Dokumente aus.

```java
  // für OpenOffice und Microsoft Office gemeinsame Einstellung
  // (Boolean.TRUE ist bereits die Standardbelegung und muss eigentlich
  // nicht explizit gesetzt werden)
  docFactory.setProperty(OfficeProperty.ERR_ON_MISSING_VAL, Boolean.TRUE);
 
  // Selbst wenn die aktuell hinter docFactory stehende Implementierung für
  // OpenOffice ist, können für den Fall auch die Microsoft Office
  // Einstellungen hinterlegt/ eingerichtet werden. Jede Implementierung
  // achtet dabei lediglich auf die eigenen Einstellungen.
  // (dieses Beispiel ist ebenfalls eine bereits bestehende Voreinstellung)
  docFactory.setProperty(MicrosoftProperty.INS_HARD_PAGE_BREAKS, Boolean.TRUE);
```

Alle Einstellungen sollten einmalig nach dem Erstellen der Factory eingerichtet werden. Für 
Dokumente bei denen die Einstellungen abweichen, sollte eine eigene Factory verwendet werden. 
Nach der Konfiguration können alle Dokumente vom Typ der Implementierung der Factory geöffnet werden.

```java
  Path invoiceDocumentPath = Paths.get("invoice-template.any");
  Path invoiceOutput = Paths.get("INVOICE-2015-SCHMIDT.DOCX");
 
  DataPage reportData = .... // siehe 'model' Package
 
  OfficeDocumentFactory docFactory = OfficeDocumentFactory
          .newInstanceByDocumentType(invoiceDocumentPath)
          .orElseThrow(IllegalStateException::new);
 
  OfficeDocument invoiceDoc = docFactory.openDocument(invoiceDocumentPath);
  invoiceDoc.generate(reportData, ResultFactory.toFile(invoiceOutput));
```

## [Spezifisch] Apache OpenOffice und LibreOffice
_Platzhalter_ sind in OpenOffice "Varibalen" vom Typ "Benutzerdefinierte Felder" (Variables "User 
Fields") und findet sich im Dialog "Felder", unter dem Tab "Variablen". Die einfachen (unbenannten!)
Platzhalter von OpenOffice werden nicht unterstützt. Jeder Platzhalter (Benutzerdefiniertes Feld) 
muss in Grußbuchstaben geschrieben sein und kann (optional) einen vordefinierten Wert enthalten. 
Der Typ wird beim Ersetzungsvorgang ignoriert und zu Text umgewandelt, ebenso wie die daraus 
resultierenden Formatierungsangaben - z.B. das Datumsformat. Wenn in der Factory eingerichtet ist, 
das fehlende Platzhalter ignoriert werden sollen, werden diese Platzhalter nicht mit dem Wert 
belassen sondern aus dem Dokument entfernt. Die umgebende Schriftformatierung von Platzhaltern 
wird im Ausgabedokument beibehalten. 

_Seitenumbrüche_ sind in OpenOffice nie hart-kodiert (im Gegensatz zu Microsoft Office) und ergeben sich lediglich aus den Formatierungsangaben von Absätzen. Soll nach jeder DataPage ein Seitenumbruch erfolgen, muss im aller ersten Absatz des Dokumentes eingestellt sein, das (Paragraph -> Text Flow -> Breaks) vor dem Absatz immer ein Seitenumbruch zu erfolgen hat. Ansonsten werden alle DataPages nur hintereinander angefügt ohne gewünschten Seitenumbruch.

_Tabellenbezeichung_ wird in OpenOffice direkt in den Eigenschaften der Tabelle unter "Name" hinterlegt. Der Name muss in Großbuchstaben eingetragen werden.

_Kopf- und Fußzeilen_ werden beim Ersetzungsvorgang nicht berücksichtigt und sollten auch keine Platzhalter enthalten.

Es werden ODF Dateien ab Version 1.1 unterstützt; nicht zu verwechseln mit der OpenOffice Versionummerierung.

## [Spezifisch] Microsoft Office
_Platzhalter_ sind Feldbefehle und ensprechen den normalen Platzhaltern in Word (MERGEFIELD). Die Groß- und Kleinschreibung ist dabei unrelevant, sollte aber am Besten groß sein. Ein mögliches MERGEFIELD ohne Beispieltext sieht wie folgt aus:

```
  { MERGEFIELD RECHNUNGSNR \* MERGEFORMAT }
```

Die geschweiften Klammern sollen dabei die Feld-Klammern von Word repräsentieren. Alternativ können auch Dokumenten-Variablen als Platzhalter missbraucht werden.

```
  { DOCVARIABLE RECHNUNGSNR }
```

Dokumentvariablen sollten dann jedoch keinerlei Verwendung in VBA (Visual Basic for Applications) Makros finden. Letzte Möglichkeit ist die Angabe eines unbekannten Feld-Befehls der nur den Namen des Platzhalters trägt. Diese Form sollte vermieden werden und ist nur implementiert um Fehlanwendung zu tolerieren. Ein vollständiges MERGEFIELD ist grundsätzlich die beste Option! Erweiterte MERGEFIELD Formatierungen/ Parameter (Datumsformatierungen, Ersatzwerte, ...) werden ignoriert und entfernt. Platzhalter übernehmen die umgebende Formatierung; im Feldbefehl selbst angegebene Formatierung, was eigentlich untypisch ist und nur durch spezielle Word-Kenntnisse möglich ist, werden ignoriert und aus dem Dokument entfernt.

_Seitenumbrüche_ können weich (per Absatzformatierung) oder hart (per Einstellung nach jeder DataPage) in Word gesetzt werden. Harte Seitenumbrüche sollten (wenn Seitenumbrüche erwünscht sind) bevorzugt werden.

_Tabellenbezeichner_ können in Word nicht direkt vergeben werden. Um einer Tabelle eine Bezeichnung/ Name zu vergeben, muss in der ersten Zelle ein unsichtbarer Textmarker hinterlegt werden, dessen Name in Großbuchstaben den Namen der Tabelle markiert.

_Kopf- und Fußzeilen_ werden in Word nicht ersetzt und sollten maximal Word-bekannte Feldbefehle enthalten.

Word-Dokumente ab Version 2007 im Open XML Document Format (DOCX) werden unterstützt.

# Daten-Modell - com.mz.solutions.office.model
Kurzlebiges Daten-Modell zur Strukturierung und späteren Befüllung von Office Dokumenten/Reports.

_Grundlegend:_
Das zur Befüllung von Dokumenten verwendete Daten-Modell ist ein kurzlebiges Modell das nur dem Zweck der Strukturierung dient. Langfristige Datenhaltung und Modellierung stehen dabei nicht im Vordergrund. Alle Modell-Klassen implementieren lediglich Funktionalität zum Hinzufügen von Bezeichner-Werte- Paaren, Tabellen und Tabellenzeilen. Eine Mutation dieser (Entfernen, Neusortieren, ...) ist nicht vorgesehen und sollte im eigenen Domain-Modell bereits erfolgt sein. Datenmodell können manuell zusammengestellt werden oder aus einer externen Quelle bezogen werden. Die Unter-Packages json und xml dienen als Implementierung zum Laden von vorbereiteten externen Datenmodellen. Die Serialisierung (und damit auch das Laden und Speichern) können alternativ ebenfalls verwendet werden.

_Wurzelelement DataPage des Dokumenten-Objekt-Modells:_
Jedes Dokument (oder jede Serie von Dokumenten/Seiten) beginnen immer mit einer Instanz von DataPage. Eine Instanz von DataPage entspricht immer einem einzelnen Dokumentes oder einer Seite oder mehrerer Seiten. Die entsprechende Interpretation ist dabei abhängig, wie das Datenmodell später übergeben wird.

_Bezeichner-Werte-Paare / Untermengen (Tabellen / Tabellenzeilen):_
Die unterschiedlichen Modell-Klassen übernehmen unterschiedliche Einträge und Strukturen auf. Die folgende Auflistung sollte dabei verdeutlichen wie die Struktur von Dokumenten ist.

```
 
  x.add(y)     | DataPage    DataTable   DataTableRow    DataValue
  -------------+--------------------------------------------------------------
  DataPage     | NEIN        JA          NEIN            JA
  DataTable    | NEIN        NEIN        JA              JA
  DataTableRow | NEIN        JA          NEIN            JA*
  DataValue    | NEIN        NEIN        NEIN            NEIN
 
  * Einfache DataValue's in DataTableRow, werden zum Ersetzen der Kopf und
    der Fußzeile verwendet; nicht zum Ersetzen von Platzhalter in den Zeilen.
``` 
 
Dabei ist zu beachten, dass DataTable und DataValue benannte Objekte sind und entsprechend eine Bezeichnung besitzen. Jede Instanz von DataValue besitzt einen Bezeichner (den Platzhalter) und jede Tabelle DataTable besitzt eine unsichtbare Tabellenbezeichnung die entweder direkt von Office als Name angegeben wird (so bei Open-Office) oder indirekt per Textmarker in der ersten Tabellenzelle versteckt angegeben wird (so bei Microsoft Office) da keine offizielle Tabellenbenennung möglich ist. Zu den genauen Unterschieden im Umgang mit Platzhaltern und Tabellennamen sollte die Package-Dokumentation von com.mz.solutions.office herangezogen werden.

__Beispiel:__

```java
  // Angaben im Hauptdokument
  DataPage invoiceDocument = new DataPage();
  invoiceDocument.addValue(new DataValue("Nachname", "Mustermann"));
  invoiceDocument.addValue(new DataValue("Vorname", "Max"));
  invoiceDocument.addValue(new DataValue("ReDatum", "2015-01-01"));
 
  // Tabelle mit den Rechnungsposten
  DataTable invoiceItems = new DataTable("Rechnungsposten");
 
  DataTableRow invItemRow1 = new DataTableRow();
  invItemRow1.addValue(new DataValue("PostenNr", "1"));
  invItemRow1.addValue(new DataValue("Artikel", "Pepsi Cola 1L"));
  invItemRow1.addValue(new DataValue("Preis", "1.70"))
 
  DataTableRow invItemRow2 = new DataTableRow();
  invItemRow2.addValues(   // Es gibt auch Vereinfachungen
          new DataValue("PostenNr", "2"),
          new DataValue("Artikel", "Kondensmilch"),
          new DataValue("Preis", "0.65"));
 
  // Hinzufügen der Zeilen; Reichenfolge ist relevant
  invoiceItems.addTableRow(invItemRow1);
  invoiceItems.addTableRow(invItemRow2);
 
  // Tabelle dem Dokument/ der Seite hinzufügen
  invoiceDocument.addTable(invoiceItems);
 
  // ...
 
```

# Schreiben der Ausgabe-Dokumente - com.mz.solutions.office.result
Der API wird eigentlich nicht direkt gesagt wo die Datei gespeichert werden soll. Wir haben dies,
da die Bibliothek auch server-seitig eingesetzt wird, soweit es geht umgangen.

Ziel und Art der Speicherung können selbst implementiert werden, indem die Schnittstelle Result implementiert wird. Alternativ können vordefinierte Implementierungen genutzt werden.

```java
  // Nutzen von fertigen Implementierungen
  Result resultToFile = ResultFactory.toFile(Paths.get("output.docx"));
 
  // oder per Lambda-Ausdruck
  Result resultToLambda = (data) -> System.out.println(Arrays.toString(data));
 
  // oder klassisch
  Result resultToInnerClass = new Result() {
      public void writeResult(byte[] dataToWrite) throws IOException {
          System.out.println(Arrays.toString(dataToWrite));
      };
  };
 
```













