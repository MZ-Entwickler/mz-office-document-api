/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2019,   Moritz Riebe     (moritz.riebe@mz-solutions.de)
 *                       Andreas Zaschka  (andreas.zaschka@mz-solutions.de)
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Affero General Public License for more details.
 * 
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

/**
 * Schnittstelle zum Öffnen und Befüllen von Dokumenten/ Reports für Apache
 * OpenOffice und Microsoft Office; hier die Package-Beschreibung lesen für den Einstieg
 * zur Nutzung dieser API/Bibliothek.
 * 
 * <p><b>Einstieg:</b><br>
 * Jeder Vorgang zum Befüllen von Textdokumenten (*.odt, *.docx, *.doct) bedarf
 * eines Dokumenten-Daten-Modelles das im Unterpackage {@code 'model'} zu finden
 * ist. Nach dem Erstellen/ Zusammensetzen des Modelles oder laden aus externen
 * Quellen, kann in weiteren Schritten mit diesem Package ein beliebiges
 * Dokument befüllt werden. <br><br></p>
 * 
 * <p><b>Öffnen von Dokumenten:</b><br>
 * Alle Dokumente werden auf dem selben Weg geöffnet. Ist das zugrunde liegende
 * Office-Format {@code (ODT oder DOCX)} bereits vor dem Öffnen bekannt, kann
 * die konrekte Implementierung gewählt werden.</p>
 * <pre>
 *  // für Libre/ Apache OpenOffice
 *  OfficeDocumentFactory docFactory = OfficeDocumentFactory.newOpenOfficeInstance();
 * 
 *  // für Microsoft Office (ab 2007)
 *  OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
 * </pre>
 * <p>Ist das Dokumentformat unbekannt, können die Implementierung selbst
 * erkennen ob es sich um ihr eigenes Format handelt.</p>
 * <pre>
 *  Path anyDocument = Paths.get("unknownFormat.blob");
 *  Optional&lt;Office&gt; docFactory = OfficeDocumentFactory.newInstanceByDocumentType(anyDocument);
 * 
 *  if (docFactory.isPresent() == false) {
 *      throw new IllegalStateException("Unbekanntes Format");
 *  }
 * </pre>
 * <p>Die Wahl zur passenden Implementierung bedarf dabei keiner korrekten
 * Angabe der Dateinamenserweiterung. Die {@link OfficeDocumentFactory} und
 * deren konkrete Implementierung erfüllt immer das Dokumenten-Daten-Modell
 * vollständig.<br><br></p>
 * 
 * <p><b>Konfiguration der Factory:</b><br>
 * Mögliche Konfigurationseigenschaften die für alle Implementierungen gültig
 * sind finden sich in {@link OfficeProperty}; Implementierungs-spezifische
 * in eigenen Klassen (z.B. {@link MicrosoftProperty}). Die Factory ist nicht
 * thread-safe und eine Änderung der Konfiguration wikt sich auf alle bereits
 * geöffneten, zukünftig zu öffnenden und derzeit in Bearbeitung befindlichen
 * Dokumente aus.</p>
 * <pre>
 *  // für OpenOffice und Microsoft Office gemeinsame Einstellung
 *  // (Boolean.TRUE ist bereits die Standardbelegung und muss eigentlich
 *  // nicht explizit gesetzt werden)
 *  docFactory.setProperty(OfficeProperty.ERR_ON_MISSING_VAL, Boolean.TRUE);
 * 
 *  // Selbst wenn die aktuell hinter docFactory stehende Implementierung für
 *  // OpenOffice ist, können für den Fall auch die Microsoft Office
 *  // Einstellungen hinterlegt/ eingerichtet werden. Jede Implementierung
 *  // achtet dabei lediglich auf die eigenen Einstellungen.
 *  // (dieses Beispiel ist ebenfalls eine bereits bestehende Voreinstellung)
 *  docFactory.setProperty(MicrosoftProperty.INS_HARD_PAGE_BREAKS, Boolean.TRUE);
 * </pre>
 * <p>Alle Einstellungen sollten einmalig nach dem Erstellen der Factory
 * eingerichtet werden. Für Dokumente bei denen die Einstellungen abweichen,
 * sollte eine eigene Factory verwendet werden. Nach der Konfiguration können
 * alle Dokumente vom Typ der Implementierung der Factory geöffnet werden.</p>
 * <pre>
 *  Path invoiceDocumentPath = Paths.get("invoice-template.any");
 *  Path invoiceOutput = Paths.get("INVOICE-2015-SCHMIDT.DOCX");
 * 
 *  DataPage reportData = .... // siehe 'model' Package
 * 
 *  OfficeDocumentFactory docFactory = OfficeDocumentFactory
 *          .newInstanceByDocumentType(invoiceDocumentPath)
 *          .orElseThrow(IllegalStateException::new);
 * 
 *  OfficeDocument invoiceDoc = docFactory.openDocument(invoiceDocumentPath);
 *  invoiceDoc.generate(reportData, ResultFactory.toFile(invoiceOutput));
 * </pre>
 * 
 * <p><b>[Spezifisch] Apache OpenOffice und LibreOffice:</b><br>
 * <i>Platzhalter</i> sind in OpenOffice "Varibalen" vom Typ "Benutzerdefinierte
 * Felder"  (Variables "User Fields") und findet sich im Dialog "Felder", unter
 * dem Tab "Variablen". Die einfachen (unbenannten!) Platzhalter von OpenOffice
 * werden nicht unterstützt. Jeder Platzhalter (Benutzerdefiniertes Feld) muss
 * in Grußbuchstaben geschrieben sein und kann (optional) einen vordefinierten
 * Wert enthalten. Der Typ wird beim Ersetzungsvorgang ignoriert und zu Text
 * umgewandelt, ebenso wie die daraus resultierenden Formatierungsangaben -
 * z.B. das Datumsformat. Wenn in der Factory eingerichtet ist, das fehlende
 * Platzhalter ignoriert werden sollen, werden diese Platzhalter nicht mit dem
 * Wert belassen sondern aus dem Dokument entfernt. Die <u>umgebende</u>
 * Schriftformatierung von Platzhaltern wird im Ausgabedokument beibehalten.
 * <br><br>
<i>Seitenumbrüche</i> sind in OpenOffice nie hart-kodiert (im Gegensatz zu
 * Microsoft Office) und ergeben sich lediglich aus den Formatierungsangaben
 * von Absätzen. Soll nach jeder {@link com.mz.solutions.office.model.DataPage}
 * ein Seitenumbruch erfolgen, muss im aller ersten Absatz des Dokumentes 
 * eingestellt sein, das (Paragraph -&gt; Text Flow -&gt; Breaks) vor dem Absatz
 * immer ein Seitenumbruch zu erfolgen hat. Ansonsten werden alle 
 * {@link com.mz.solutions.office.model.DataPage}s nur hintereinander angefügt 
 * ohne gewünschten Seitenumbruch.<br><br>
 * <i>Tabellenbezeichung</i> wird in OpenOffice direkt in den Eigenschaften
 * der Tabelle unter "Name" hinterlegt. Der Name muss in Großbuchstaben
 * eingetragen werden.<br><br>
 * <i>Kopf- und Fußzeilen</i> werden beim normalen Ersetzungsvorgang nicht berücksichtigt.
 * Mit Dokumenten-Anweisungen können Kopf- und Fußzeilen bei ODT Dokumenten ersetzt werden.
 * Siehe dazu Abschnitt Dokumenten-Anweisungen.<br><br>
 * Es werden ODF Dateien ab Version 1.1 unterstützt; nicht zu verwechseln mit
 * der OpenOffice Versionummerierung.<br><br></p>
 * 
 * <p><b>[Spezifisch] Microsoft Office:</b><br>
 * <i>Platzhalter</i> sind Feldbefehle und ensprechen den normalen Platzhaltern
 * in Word ({@code MERGEFIELD}). Die Groß- und Kleinschreibung ist dabei
 * unrelevant, sollte aber am Besten groß sein. Ein mögliches {@code MERGEFIELD}
 * ohne Beispieltext sieht wie folgt aus:</p><pre>
 *  { MERGEFIELD RECHNUNGSNR \* MERGEFORMAT }
 * </pre><p>Die geschweiften Klammern sollen dabei die Feld-Klammern von Word
 * repräsentieren. Alternativ können auch Dokumenten-Variablen als Platzhalter
 * missbraucht werden.</p><pre>
 *  { DOCVARIABLE RECHNUNGSNR }
 * </pre><p>Dokumentvariablen sollten dann jedoch keinerlei Verwendung in
 * VBA (Visual Basic for Applications) Makros finden. Letzte Möglichkeit ist die
 * Angabe eines unbekannten Feld-Befehls der nur den Namen des Platzhalters
 * trägt. <u>Diese Form sollte vermieden werden und ist nur implementiert um
 * Fehlanwendung zu tolerieren.</u> Ein vollständiges {@code MERGEFIELD} ist
 * grundsätzlich die beste Option! Erweiterte {@code MERGEFIELD} Formatierungen/
 * Parameter (Datumsformatierungen, Ersatzwerte, ...) werden ignoriert und
 * entfernt. Platzhalter übernehmen die <u>umgebende</u> Formatierung; im
 * Feldbefehl selbst angegebene Formatierung, was eigentlich untypisch ist
 * und nur durch spezielle Word-Kenntnisse möglich ist, werden ignoriert und
 * aus dem Dokument entfernt.<br><br>
 * <i>Seitenumbrüche</i> können weich (per Absatzformatierung) oder hart (per
 * Einstellung nach jeder {@link com.mz.solutions.office.model.DataPage}) in
 * Word gesetzt werden. Harte Seitenumbrüche sollten (wenn Seitenumbrüche
 * erwünscht sind) bevorzugt werden.<br><br>
 * <i>Tabellenbezeichner</i> können in Word nicht direkt vergeben werden. Um
 * einer Tabelle eine Bezeichnung/ Name zu vergeben, muss in der ersten Zelle
 * ein unsichtbarer Textmarker hinterlegt werden, dessen Name in Großbuchstaben
 * den Namen der Tabelle markiert.<br><br>
 * <i>Kopf- und Fußzeilen</i> werden in Word nicht ersetzt und sollten maximal
 * Word-bekannte Feldbefehle enthalten.<br><br>
 * Word-Dokumente ab Version 2007 im Open XML Document Format ({@code DOCX})
 * werden unterstützt.</p>
 * 
 * <p><b>Umgang mit dem Einsetzen/Ersetzen von Bildern in Dokumenten:</b><br>
 * Bilder können in Vorlage-Dokumenten eingesetzt sowie durch andere ersetzt werden. Bei
 * Text-Platzhaltern (MergeFields bei Microsoft, User-Def-Fields bei Libre/Openoffice), wird bei
 * einem Bild-Wert {@link com.mz.solutions.office.model.images.ImageValue} an jene Stelle das
 * als {@link com.mz.solutions.office.model.images.ImageResource} geladene Bild eingesetzt unter
 * Verwendung der angegebenen Abmaße aus dem Bild-Wert.<br>
 * Bestehende Bilder können ersetzt/ausgetauscht werden und, wenn gewünscht, deren bestehenden
 * Abmaße in der Vorlage mit eigenen überschrieben/ersetzt werden. Bild-Platzhalter, also in der
 * Vorlage bereits existierende Bilder, werden als Platzhalter erkannt, wenn dem Bild in der
 * Vorlage in den Eigenschaften [Titel, Name, Beschreibung, Alt-Text] ein bekannter Platzhalter
 * mit Bild-Wert angegeben wurde. Genaueres ist den folgenden Klassen zu entnehmen:</p>
 * 
 * <pre>
 *  {@link com.mz.solutions.office.model.images.ImageResource}
 *  Bild-Datei/-Resource (Bild als Byte-Array mit Angabe des Formates)
 * 
 *  {@link com.mz.solutions.office.model.images.ImageValue}
 *  Bild-Wert (Resource) mit weiteren Angaben wie Titel (optional), Beschreibung (optional)
 *  und anzuwendende Abmaße.
 * </pre>
 * 
 * <p>Ein Bild-Wert ({@link com.mz.solutions.office.model.images.ImageValue}) besitzt eine
 * zugeordnete Bild-Resource ({@link com.mz.solutions.office.model.images.ImageResource}). Eine
 * Bild-Resource kann mehrfach/gleichzeitig in mehreren Bild-Werten verwendet werden.
 * Das Wiederverwenden von Bild-Resourcen führt zu deutlich kleineren Ergebnis-Dokumenten. Jene
 * Bild-Resource wird dann nur einmalig im Ergebnis-Dokument eingebettet.</p>
 * 
 * <pre>
 *  // Create or load Image-Resources. Try to reuse resources to reduce the file size of the
 *  // result documents. Internally image resources cache the file content.
 *  ImageResource imageData1 = ImageResource.loadImage(
 *          Paths.get("image_1.png"), StandardImageResourceType.PNG);
 * 
 *  ImageResource imageData2 = ImageResource.loadImage(
 *          Paths.get("image_2.bmp"), StandardImageResourceType.BMP);
 * 
 *  ImageValue image1Small = new ImageValue(imageData1)
 *          .setDimension(0.5D, 0.5D, UnitOfLength.CENTIMETERS)     // default 3cm x 1cm
 *          .setTitle("Image Title")                                // optional
 *          .setDescription("Alternative Text Description");        // optional
 * 
 *  ImageValue image1Large = new ImageValue(imageData1) // same image as image1Small (sharing res.)
 *          .setDimension(15, 15, UnitOfLength.CENTIMETERS);
 * 
 *  ImageValue image2 = new ImageValue(imageData2)
 *          .setDimension(40, 15, UnitOfLength.MILLIMETERS)
 *          .setOverrideDimension(true);
 * 
 *  // Assigning ImageValue's to DataValue's
 *  final DataPage page = new DataPage();
 * 
 *  page.addValue(new DataValue("IMAGE_1_SMALL", image1Small));
 *  page.addValue(new DataValue("IMAGE_1_LARGE", image1Large));
 *  page.addValue(new DataValue("IMAGE_2", image2));
 *  page.addValue(new DataValue("IMAGE_B", image2)); // ImageValue's are reusable
 * </pre>
 * 
 * <p><b>Dokument-Anweisungen mit
 * {@link com.mz.solutions.office.instruction.DocumentProcessingInstruction} übergeben:</b><br>
 * Dem Ersetzungs-Vorgang können weitere Anweisungen/Callbacks mit übergeben werden. Derzeit
 * mögliche Anweisungen ist das Abfangen (oder gezielte Laden) von Dokumenten-Teilen (also XML
 * Dateien im ZIP Container) und der Bearbeitung des XML-Baumes vor und/oder nach Ausführung des
 * Ersetzungs-Vorganges.<br>
 * Bei LibreOffice/Apache-OpenOffice können dazu bei ODT Dateien die Kopf- und Fußzeilen im
 * Ersetzungs-Prozess mit einbezogen werden.<br>
 * Alle Anweisungen können erstellt werden über die vereinfachten Factory-Methoden in
 * {@link com.mz.solutions.office.instruction.DocumentProcessingInstruction} oder händisch durch
 * Implementieren der jeweiligen Klassen.</p>
 * 
 * <p>Kopf- und Fußzeilen werden (derzeit) nur bei {@code ODT} Dokumente unterstützt.</p>
 * <pre>
 *  // Header and Footer in ODT Documents (Header and Footer in MS Word Documents are not supported)
 *  final OfficeDocument anyDocument = ...
 * 
 *  final DataPage documentData = ...
 *  final DataPage headerData = ...     // Header und Footer replacement only for ODT-Files
 *  final DataPage footerData = ...
 * 
 *  anyDocument.generate(documentData, ResultFactory.toFile(invoiceOutput),
 *          DocumentProcessingInstruction.replaceHeaderWith(headerData),
 *          DocumentProcessingInstruction.replaceFooterWith(footerData));
 * </pre>
 * 
 * <p>Document-Interceptors werden bei beiden Office-Implementierungen unterstützt.</p>
 * <pre>
 *  final DocumentInterceptorFunction interceptorFunction = ...
 *  final DocumentInterceptorFunction changeCustomXml = (DocumentInterceptionContext context) -&gt; {
 *      final Document xmlDocument = context.getXmlDocument();
 *      final NodeList styleNodes = xmlDocument.getElementsByTagName("custXml:customers");
 * 
 *      // add/remove/change XML document
 *      // 'context' should contain all data and access you will need
 *  };
 * 
 *  anyDocument.generate(documentData, ResultFactory.toFile(invoiceOutput),
 *          // Intercept main document part (document body)
 *          DocumentProcessingInstruction.interceptDocumentBody(
 *                  DocumentInterceptorType.BEFORE_GENERATION,  // invoke interceptor before
 *                  interceptorFunction, // change low level document function (Callback-Method)
 *                  dataMapForInterceptorFunctionHere), // data is optional
 *          // Intercept styles part of this document, maybe to change font-scaling afterwards
 *          DocumentProcessingInstruction.interceptDocumentStylesAfter(
 *                  interceptorFunction), // no data for this callback function
 *          // let us change the Custom XML Document Part (only MS Word) und fill with our data
 *          DocumentProcessingInstruction.interceptXmlDocumentPart(
 *                  "word/custom/custPropItem.xml", // our Custom XML data
 *                  DocumentInterceptorType.AFTER_GENERATION, // before or after doesn't matter
 *                  changeCustomXml));
 *          
 * </pre>
 */
package com.mz.solutions.office;
