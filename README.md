# MZ Office Document API
![](https://img.shields.io/badge/maven-0.8.0-green)  ![Java CI with Maven](https://github.com/MZ-Entwickler/mz-office-document-api/workflows/Java%20CI%20with%20Maven/badge.svg?branch=master)  
Library to fill out Libre-/OpenOffice (odt) and Microsoft Word Documents (docx; starting with Word 2007) containing place holders and nested Tables.
The data model is independent of the resulting document format.
The library is not intended to create documents from scratch, instead to work with templates with placeholders.

By using an abstract model layer (`com.mz.solutions.office.model`),
both formats (docx/odt) can be used by the same code base.

For Microsoft Word there are some extensions to use Custom XML Parts.

## Lizenz
Affero General Public License Version 3 (AGPL v3)

## Abhängigkeiten
- _Runtime_: No dependencies at runtime needed
- _Compile_: JSR 305 (Annotations for Software Defect Detection)
- _Test_: JUnit 5

The project needs Java 13

# Examples
In the directory `examples` there are some example projects. Further information can be found in the java doc.

# Einstieg - com.mz.solutions.office
Jeder Vorgang zum Befüllen von Textdokumenten (*.odt, *.docx, *.doct) bedarf eines
Dokumenten-Daten-Modelles das im Unterpackage 'model' zu finden ist. Nach dem Erstellen/
Zusammensetzen des Modelles oder laden aus externen Quellen, kann in weiteren Schritten mit 
diesem Package ein beliebiges Dokument befüllt werden.

## Dokumente öffnen
Alle Dokumente werden auf dem selben Weg geöffnet. Ist das zugrunde liegende Office-Format 
(ODT oder DOCX) bereits vor dem Öffnen bekannt, kann die konrekte Implementierung gewählt werden.

```java
  // for Libre/ Apache OpenOffice
  OfficeDocumentFactory docFactory = OfficeDocumentFactory.newOpenOfficeInstance();
 
  // for Microsoft Office (starting with 2007)
  OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
```

Ist das Dokumentformat unbekannt, können die Implementierung selbst erkennen ob es sich um ihr
eigenes Format handelt.

```java
  Path anyDocument = Paths.get("unknownFormat.blob");
  Optional<Office> docFactory = OfficeDocumentFactory.newInstanceByDocumentType(anyDocument);
 
  if (docFactory.isPresent() == false) {
      throw new IllegalStateException("unknown format");
  }
```

Die Wahl zur passenden Implementierung bedarf dabei keiner korrekten Angabe der 
Dateinamenserweiterung. Die OfficeDocumentFactory und deren konkrete Implementierung erfüllt 
immer das Dokumenten-Daten-Modell vollständig.

## Konfiguration der Factory
Mögliche Konfigurationseigenschaften die für alle Implementierungen gültig sind finden sich in 
`OfficeProperty`; Implementierungs-spezifische in eigenen Klassen (z.B. `MicrosoftProperty`). Die 
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

_Seitenumbrüche_ sind in OpenOffice nie hart-kodiert (im Gegensatz zu Microsoft Office) und ergeben sich lediglich aus den Formatierungsangaben von Absätzen. Soll nach jeder `DataPage` ein Seitenumbruch erfolgen, muss im aller ersten Absatz des Dokumentes eingestellt sein, das (Paragraph -> Text Flow -> Breaks) vor dem Absatz immer ein Seitenumbruch zu erfolgen hat. Ansonsten werden alle `DataPage`s nur hintereinander angefügt ohne gewünschten Seitenumbruch.

_Tabellenbezeichung_ wird in OpenOffice direkt in den Eigenschaften der Tabelle unter "Name" hinterlegt. Der Name muss in Großbuchstaben eingetragen werden.

_Kopf- und Fußzeilen_ werden beim normalen Ersetzungsvorgang nicht berücksichtigt. Mit Dokumenten-Anweisungen können Kopf- und Fußzeilen bei ODT Dokumenten ersetzt werden. Siehe dazu Abschnitt Dokumenten-Anweisungen.

Es werden ODF Dateien ab Version 1.1 unterstützt; nicht zu verwechseln mit der OpenOffice Versionummerierung.

## [Spezifisch] Microsoft Office
_Platzhalter_ sind Feldbefehle und ensprechen den normalen Platzhaltern in Word (`MERGEFIELD`). Die Groß- und Kleinschreibung ist dabei unrelevant, sollte aber am Besten groß sein. Ein mögliches `MERGEFIELD` ohne Beispieltext sieht wie folgt aus:

```
  { MERGEFIELD RECHNUNGSNR \* MERGEFORMAT }
```

Die geschweiften Klammern sollen dabei die Feld-Klammern von Word repräsentieren. Alternativ können auch Dokumenten-Variablen als Platzhalter missbraucht werden. Wichtig: Nicht in Word einfach geschweifte Klammern schreiben! Die geschweiften Klammern werden mit `STRG+F9` erzeugt und sind in Word besondere Feldbefehle.

```
  { DOCVARIABLE RECHNUNGSNR }
```

Dokumentvariablen sollten dann jedoch keinerlei Verwendung in VBA (Visual Basic for Applications) Makros finden. Letzte Möglichkeit ist die Angabe eines unbekannten Feld-Befehls der nur den Namen des Platzhalters trägt. Diese Form sollte vermieden werden und ist nur implementiert um Fehlanwendung zu tolerieren. Ein vollständiges `MERGEFIELD` ist grundsätzlich die beste Option! Erweiterte `MERGEFIELD` Formatierungen/ Parameter (Datumsformatierungen, Ersatzwerte, ...) werden ignoriert und entfernt. Platzhalter übernehmen die umgebende Formatierung; im Feldbefehl selbst angegebene Formatierung, was eigentlich untypisch ist und nur durch spezielle Word-Kenntnisse möglich ist, werden ignoriert und aus dem Dokument entfernt.

_Seitenumbrüche_ können weich (per Absatzformatierung) oder hart (per Einstellung nach jeder `DataPage`) in Word gesetzt werden. Harte Seitenumbrüche sollten (wenn Seitenumbrüche erwünscht sind) bevorzugt werden.

_Tabellenbezeichner_ können in Word nicht direkt vergeben werden. Um einer Tabelle eine Bezeichnung/ Name zu vergeben, muss in der ersten Zelle ein unsichtbarer Textmarker hinterlegt werden, dessen Name in Großbuchstaben den Namen der Tabelle markiert.

_Kopf- und Fußzeilen_ werden in Word nicht ersetzt und sollten maximal Word-bekannte Feldbefehle enthalten.

Word-Dokumente ab Version 2007 im Open XML Document Format (DOCX) werden unterstützt.

# Daten-Modell - com.mz.solutions.office.model
Kurzlebiges Daten-Modell zur Strukturierung und späteren Befüllung von Office Dokumenten/Reports.

_Grundlegend:_
Das zur Befüllung von Dokumenten verwendete Daten-Modell ist ein kurzlebiges Modell das nur dem Zweck der Strukturierung dient. Langfristige Datenhaltung und Modellierung stehen dabei nicht im Vordergrund. Alle Modell-Klassen implementieren lediglich Funktionalität zum Hinzufügen von Bezeichner-Werte- Paaren, Tabellen und Tabellenzeilen. Eine Mutation dieser (Entfernen, Neusortieren, ...) ist nicht vorgesehen und sollte im eigenen Domain-Modell bereits erfolgt sein. Datenmodell können manuell zusammengestellt werden oder aus einer externen Quelle bezogen werden. Die Unter-Packages `json` und `xml` (in der GitHub Version noch nicht verfügbar, da noch im Beta-Status) dienen als Implementierung zum Laden von vorbereiteten externen Datenmodellen. Die Serialisierung (und damit auch das Laden und Speichern) können alternativ ebenfalls verwendet werden.

_Wurzelelement DataPage des Dokumenten-Objekt-Modells:_
Jedes Dokument (oder jede Serie von Dokumenten/Seiten) beginnen immer mit einer Instanz von `DataPage`. Eine Instanz von `DataPage` entspricht immer einem einzelnen Dokumentes oder einer Seite oder mehrerer Seiten. Die entsprechende Interpretation ist dabei abhängig, wie das Datenmodell später übergeben wird.

_Bezeichner-Werte-Paare / Untermengen (Tabellen / Tabellenzeilen):_
Die unterschiedlichen Modell-Klassen übernehmen unterschiedliche Einträge und Strukturen auf. Die folgende Auflistung sollte dabei verdeutlichen wie die Struktur von Dokumenten ist.

```
 
  x.add(y)     | DataPage    DataTable   DataTableRow    DataValue
  -------------+--------------------------------------------------------------
  DataPage     | NEIN        JA          NEIN            JA
  DataTable    | NEIN        NEIN        JA              JA*
  DataTableRow | NEIN        JA          NEIN            JA
  DataValue    | NEIN        NEIN        NEIN            NEIN
 
  * Einfache DataValue's in DataTable, werden zum Ersetzen der Kopf und
    der Fußzeile in der Tabelle verwendet; nicht zum Ersetzen von Platzhalter in den Zeilen.
``` 
 
Dabei ist zu beachten, dass `DataTable` und `DataValue` benannte Objekte sind und entsprechend eine Bezeichnung besitzen. Jede Instanz von `DataValue` besitzt einen Bezeichner (den Platzhalter) und jede Tabelle `DataTable` besitzt eine unsichtbare Tabellenbezeichnung die entweder direkt von Office als Name angegeben wird (so bei Open-Office) oder indirekt per Textmarker in der ersten Tabellenzelle versteckt angegeben wird (so bei Microsoft Office) da keine offizielle Tabellenbenennung möglich ist. Zu den genauen Unterschieden im Umgang mit Platzhaltern und Tabellennamen sollte die Package-Dokumentation von `com.mz.solutions.office` herangezogen werden.

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

# Umgang mit dem Einsetzen/Ersetzen von Bildern in Dokumenten
Bilder können in Vorlage-Dokumenten eingesetzt sowie durch andere ersetzt werden. Bei
Text-Platzhaltern (MergeFields bei Microsoft, User-Def-Fields bei Libre/Openoffice), wird bei
einem Bild-Wert `com.mz.solutions.office.model.images.ImageValue` an jene Stelle das
als `com.mz.solutions.office.model.images.ImageResource` geladene Bild eingesetzt unter
Verwendung der angegebenen Abmaße aus dem Bild-Wert.

Bestehende Bilder können ersetzt/ausgetauscht werden und, wenn gewünscht, deren bestehenden
Abmaße in der Vorlage mit eigenen überschrieben/ersetzt werden. Bild-Platzhalter, also in der
Vorlage bereits existierende Bilder, werden als Platzhalter erkannt, wenn dem Bild in der
Vorlage in den Eigenschaften (Titel, Name, Beschreibung, Alt-Text) ein bekannter Platzhalter
mit Bild-Wert angegeben wurde. Genaueres ist den folgenden Klassen zu entnehmen:

`com.mz.solutions.office.model.images.ImageResource`
 Bild-Datei/-Resource (Bild als Byte-Array mit Angabe des Formates)

 `com.mz.solutions.office.model.images.ImageValue`
 Bild-Wert (Resource) mit weiteren Angaben wie Titel (optional), Beschreibung (optional)
 und anzuwendende Abmaße.

Ein Bild-Wert (`com.mz.solutions.office.model.images.ImageValue`) besitzt eine
zugeordnete Bild-Resource (`com.mz.solutions.office.model.images.ImageResource`). Eine
Bild-Resource kann mehrfach/gleichzeitig in mehreren Bild-Werten verwendet werden.
Das Wiederverwenden von Bild-Resourcen führt zu deutlich kleineren Ergebnis-Dokumenten. Jene
Bild-Resource wird dann nur einmalig im Ergebnis-Dokument eingebettet.

```java
 // Create or load Image-Resources. Try to reuse resources to reduce the file size of the
 // result documents. Internally image resources cache the file content.
 ImageResource imageData1 = ImageResource.loadImage(
         Paths.get("image_1.png"), StandardImageResourceType.PNG);

 ImageResource imageData2 = ImageResource.loadImage(
         Paths.get("image_2.bmp"), StandardImageResourceType.BMP);

 ImageValue image1Small = new ImageValue(imageData1)
         .setDimension(0.5D, 0.5D, UnitOfLength.CENTIMETERS)     // default 3cm x 1cm
         .setTitle("Image Title")                                // optional
         .setDescription("Alternative Text Description");        // optional

 ImageValue image1Large = new ImageValue(imageData1) // same image as image1Small (sharing res.)
         .setDimension(15, 15, UnitOfLength.CENTIMETERS);

 ImageValue image2 = new ImageValue(imageData2)
         .setDimension(40, 15, UnitOfLength.MILLIMETERS)
         .setOverrideDimension(true);

 // Assigning ImageValue's to DataValue's
 final DataPage page = new DataPage();

 page.addValue(new DataValue("IMAGE_1_SMALL", image1Small));
 page.addValue(new DataValue("IMAGE_1_LARGE", image1Large));
 page.addValue(new DataValue("IMAGE_2", image2));
 page.addValue(new DataValue("IMAGE_B", image2)); // ImageValue's are reusable
```

# Dokument-Anweisungen mit `DocumentProcessingInstruction` übergeben
Dem Ersetzungs-Vorgang können weitere Anweisungen/Callbacks mit übergeben werden. Derzeit
mögliche Anweisungen ist das Abfangen (oder gezielte Laden) von Dokumenten-Teilen (also XML
Dateien im ZIP Container) und der Bearbeitung des XML-Baumes vor und/oder nach Ausführung des
Ersetzungs-Vorganges.

Bei LibreOffice/Apache-OpenOffice können dazu bei ODT Dateien die Kopf- und Fußzeilen im
Ersetzungs-Prozess mit einbezogen werden.

Alle Anweisungen können erstellt werden über die vereinfachten Factory-Methoden in
`com.mz.solutions.office.instruction.DocumentProcessingInstruction` oder händisch durch
Implementieren der jeweiligen Klassen.

__Kopf- und Fußzeilen werden (derzeit) nur bei `ODT` Dokumente unterstützt.__
```java
 // Header and Footer in ODT Documents (Header and Footer in MS Word Documents are not supported)
 final OfficeDocument anyDocument = ...

 final DataPage documentData = ...
 final DataPage headerData = ...     // Header und Footer replacement only for ODT-Files
 final DataPage footerData = ...

 anyDocument.generate(documentData, ResultFactory.toFile(invoiceOutput),
         DocumentProcessingInstruction.replaceHeaderWith(headerData),
         DocumentProcessingInstruction.replaceFooterWith(footerData));
```

___Document-Interceptors werden bei beiden Office-Implementierungen unterstützt.___
```java
 final DocumentInterceptorFunction interceptorFunction = ...
 final DocumentInterceptorFunction changeCustomXml = (DocumentInterceptionContext context) -> {
     final Document xmlDocument = context.getXmlDocument();
     final NodeList styleNodes = xmlDocument.getElementsByTagName("custXml:customers");

     // add/remove/change XML document
     // 'context' should contain all data and access you will need
 };

 anyDocument.generate(documentData, ResultFactory.toFile(invoiceOutput),
         // Intercept main document part (document body)
         DocumentProcessingInstruction.interceptDocumentBody(
                 DocumentInterceptorType.BEFORE_GENERATION,  // invoke interceptor before
                 interceptorFunction, // change low level document function (Callback-Method)
                 dataMapForInterceptorFunctionHere), // data is optional
         // Intercept styles part of this document, maybe to change font-scaling afterwards
         DocumentProcessingInstruction.interceptDocumentStylesAfter(
                 interceptorFunction), // no data for this callback function
         // let us change the Custom XML Document Part (only MS Word) und fill with our data
         DocumentProcessingInstruction.interceptXmlDocumentPart(
                 "word/custom/custPropItem.xml", // our Custom XML data
                 DocumentInterceptorType.AFTER_GENERATION, // before or after doesn't matter
                 changeCustomXml));
         
```






