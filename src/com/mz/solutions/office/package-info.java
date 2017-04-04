/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2017,   Moritz Riebe     (moritz.riebe@mz-solutions.de)
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
 * <i>Kopf- und Fußzeilen</i> werden beim Ersetzungsvorgang nicht berücksichtigt
 * und sollten auch keine Platzhalter enthalten.<br><br>
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
 */
package com.mz.solutions.office;
