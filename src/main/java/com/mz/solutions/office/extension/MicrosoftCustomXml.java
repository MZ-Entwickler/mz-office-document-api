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
package com.mz.solutions.office.extension;

import com.mz.solutions.office.OfficeDocumentException;
import org.w3c.dom.Document;
import org.w3c.dom.Node;

/**
 * Microsoft Office Erweiterung: Vereinfachung des Ersetzungsvorgangs bei
 * Nutzung von Custom XML in Word-Dokumenten.
 *
 * <p>Implementiert einen einfachen Zugriff auf Custom XML Parts für Word
 * Dokumente ab 2007.</p>
 *
 * <p>Der Zustand der Custom XML Parts ist unabhängig der normalen Platzhalter
 * und kann auch in Kombinationen mit jenen (soweit sinnvoll) verwendet werden.
 * Nach dem Überschreiben einzelner Parts und der Generierung des neuen
 * Dokumentes, bleibt der Zustand erhalten.</p>
 *
 * <p>Bei einem Dokument dessen Custom XML Part verändert werden sollen, aber
 * es werden keine Parts gefunden, führt das Abrufen der Erweiterung nicht
 * zu einem Fehler. Lediglich die Rückgabe von {@link #countParts()} liefert
 * {@code 0} zurück.</p>
 *
 * <pre>
 *  final Document msOfficeDocument = ...
 *  final Optional&lt;MicrosoftCustomXml&gt; opCustXmlExt =
 *          msOfficeDocument.extension(MicrosoftCustomXml.class);
 *
 *  if (opCustXmlExt.isPresent() == false) {
 *      // Format unterstützt keine Custom XML Parts.
 *      // Dies ist NICHT gleichbedeutet damit, das keine XML Parts im gegebenen
 *      // Dokument enthalten sind!
 *  }
 *
 *  final MicrosoftCustomXml custXml = opCustXmlExt.get();
 *
 *  ... custXml.overwritePartAt(...)... // Änderung der XML Inhalte
 *
 *  msOfficeDocument.generate(...) // Schreiben mit neuen Datensätzen
 *
 *  // Nach dem Schreibvorgang bleibt der Zustand von custXml erhalten und muss
 *  // für den nächsten Schreibvorgang verändert werden.
 * </pre>
 *
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 * @see <a href="https://msdn.microsoft.com/en-us/library/bb608618.aspx">
 * (MSDN) Custom XML Parts Overview</a>
 */
public abstract class MicrosoftCustomXml implements Extension {

    protected MicrosoftCustomXml() {
    }

    /**
     * Gibt die Anzahl der gefundenen Custom XML Parts zurück.
     *
     * @return Anzahl der eingebetteten (mapped) XML Parts
     */
    public abstract int countParts();

    /**
     * Gibt den Inhalt eines Parts am gegebenen Index zurück.
     *
     * @param index Index ab 0.
     * @return XML als Byte-Array
     * @throws IllegalStateException     Wenn keine Parts vorhanden sind.
     * @throws IndexOutOfBoundsException Wenn am gegebenen Index kein Custom XML Part vorhanden ist.
     * @throws OfficeDocumentException   Wenn beim Abrufen, Auslesen, etc.. ein Fehler auftritt.
     */
    public abstract byte[] partAsBytesAt(int index)
            throws OfficeDocumentException;

    /**
     * Gibt den Inhalt eines Parts als Zeichenkette zurück unter der Annahme
     * das das hinterlegte XML in UTF-8 Kodierung gespeichert ist/ wurde.
     *
     * <p>Verhält sich identisch zu {@link #partAsBytesAt(int)}</p>
     *
     * @param index Index ab 0.
     * @return Zeichenkette mit XML (angenommen wird UTF-8).
     */
    public abstract String partAsStringAt(int index)
            throws OfficeDocumentException;

    /**
     * Gibt den Inhalt eines Parts als Zeichenkette zurück in Form eines
     * XML DOM Dokumentes.
     *
     * <p>Verhält sich identisch zu {@link #partAsBytesAt(int)}.</p>
     *
     * @param index Index ab 0.
     * @return Custom XML Part als XML DOM Dokument.
     */
    public abstract Document partAsDocumentAt(int index)
            throws OfficeDocumentException;

    /**
     * Überschreibt ein bereits bestehenden Part mit eigenen Daten in Form
     * eines Byte-Arrays (welches intern XML entspricht).
     *
     * <p>Diese Methode fügt <b>keine</b> neuen Parts hinzu, sondern
     * überschreibt lediglich einen bestehenden Part im Dokument.</p>
     *
     * @param index Index ab 0.
     * @param value XML als Byte-Array
     * @throws IllegalStateException     Wenn keine Parts vorhanden sind.
     * @throws IndexOutOfBoundsException Wenn am gegebenen Index kein Custom XML Part vorhanden ist.
     */
    public abstract void overwritePartAt(int index, byte[] value);

    /**
     * Überschreibt einen bereits bestehenden Part mit eigenen XML Werten in
     * der übergebenen Zeichenkette die als UTF-8 kodiert wird.
     *
     * <p>Verhält sich identisch zu {@link #overwritePartAt(int, byte[])}.</p>
     *
     * @param index Index ab 0.
     * @param value XML als Zeichenkette; wird als UTF-8 kodiert.
     */
    public abstract void overwritePartAt(int index, String value);

    /**
     * Überschreibt einen bereits bestehenden Part durch Übergabe eines XML DOM
     * Nodes (oder Documents).
     *
     * <p>Verhält sich identisch zu {@link #overwritePartAt(int, byte[])}.</p>
     *
     * @param index Index ab 0.
     * @param value Node/Document das vor dem Speichern normalisiert wird
     *              und als UTF-8 kodiert wird.
     */
    public abstract void overwritePartAt(int index, Node value);

}
