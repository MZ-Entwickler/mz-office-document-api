/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2018,   Moritz Riebe     (moritz.riebe@mz-solutions.de)
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

import com.mz.solutions.office.model.DataValue;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Objects;
import static java.util.Objects.requireNonNull;
import java.util.Optional;
import java.util.UUID;
import javax.annotation.CheckReturnValue;
import javax.annotation.Nullable;
import static com.mz.solutions.office.resources.MessageResources.formatMessage;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * Microsoft Office Erweiterung einfügen/importieren/integrieren von externen Dokumenten an
 * Absatzsstellen in Word-Dokumenten; unterschiedlichen Formates.
 * 
 * <p>Einzufügende Teile können in Word-Dokumenten nur als ganze Absätze importiert werden und nicht
 * in oder zwischen Fließtexten. Als ganze Absätze zählen: der Absatz als solcher, das Hauptdokument
 * (der Body), Kommentarfelder und Tabellenzellen ({@code w:tc}). Innerhalb von Tabellenzellen auch,
 * soweit sich der Platzhalter dafür wieder in einem Absatz befindet.</p>
 * 
 * <p>Wird der Platzhalter in einem Fließtext gefunden, an dessen Stelle das Dokument importiert
 * werden soll, wird der Platzhalter entfernt und an der frühstmöglichen Stelle (es wird im Dokument
 * also soweit nach "oben gesprungen" bis es möglich ist) dann importiert. Dies kann (soweit man
 * sich der Einschränkung nicht bewusst ist) zu einer Verschiebung des Ersetzungsortes führen.</p>
 * 
 * <p>Um Dokumente zur deklarieren, die importiert werden sollen, kann die statische Funktion
 * {@link #insertFile(Path, ImportFormatType)} verwendet werden. Der zurückgegebene Wert
 * wird einfach einer {@link DataValue} Instanz übergeben, anstelle der Zeichenkette.</p>
 * 
 * <p>Zulässige Formate zum importieren/einfügen in Word-Dokumenten sind im Enum
 * {@link ImportFormatType} vollständig zusammengefasst und aufgezählt.</p>
 * 
 * <p>Besonderer Hinweis zur Serialisierung: Mit dieser Klasse erzeugte Erweiterte-Werte
 * ({@link ExtendedValue}'s) sind serialisierbar, <b>aber</b> diese speichern nicht den Inhalt des
 * einzufügenden Dokumentes, sondern lediglich den Pfad zum Dokument. Wird der Erweiterte-Wert
 * verwendet, serialisiert und zu einem anderen Zeitpunkt <u>oder</u> auf einer anderen Instanz
 * abgerufen, muss sichergestellt sein, dass das gewünschte Dokument dort vorhanden ist!</p>
 * 
 * <p><b>Beispiel</b></p>
 * 
 * <pre>
  //// Ausführliche Belegung unter Angabe des konkreten Formates
 
  final Path topText = Paths.get("C:\\DocumentParts\\InvoiceTopText.docx");
  final Path bottomText = Paths.get("C:\\DocumentParts\\InvoiceBottomText.docx");
 
  // Funktioniert nur, wenn der Dateiname die korrekte Dateiendung beinhaltet, ansonsten
  // müsste manuell der korrekte Formatwert vom Enum angegeben werden!
  final ImportFormatType formatDocTop = ImportFormatType.byFileExtension(topText).get();
  final ImportFormatType formatDocBottom = ImportFormatType.byFileExtension(bottomText).get();
 
  final DataPage page = new DataPage();
  final ExtendedValue extValTop = MicrosoftInsertDoc.insertFile(topText, formatDocTop);
  final ExtendedValue extValBot = MicrosoftInsertDoc.insertFile(bottomText, formatDocBottom);
 
  page.addValue(new DataValue("TopTextIns", extValTop));
  page.addValue(new DataValue("BottomTextIns", extValBottom));
 
 
  //// Alternativ ginge auch die Kurzform in der grundsätzlich davon ausgegangen wird, dass das
  //// Dokumentformat anhand des Dateinamens korrekt erkannt werden kann.
 
  final Path topText = Paths.get("C:\\DocumentParts\\InvoiceTopText.docx");
  final Path bottomText = Paths.get("C:\\DocumentParts\\InvoiceBottomText.docx");
 
  final DataPage page = new DataPage();
 
  page.addValue(new DataValue("TopTextIns", MicrosoftInsertDoc.insertFile(topText)));
  page.addValue(new DataValue("BottomTextIns", MicrosoftInsertDoc.insertFile(bottomText)));
 </pre>
 * 
 * <p><b>Empfehlung bei Verwendung:</b> Wird ein Textbaustein innerhalb eines Dokumentes mehrfach
 * verwendet, dann ist es von Vorteil für die Ladezeit und Größe der späteren Datei, das die Instanz
 * von {@link ExtendedValue} nicht mehrfach erzeugt wird, sondern nach einmaliger Erzeugung mehrfach
 * verwendet wird.</p>
 * 
 * <p>Leere Texteinsetzungen (also Platzhalter an denen kein Dokument eingefügt werden sollen)
 * können über eine leere Zeichenkette (normales {@link DataValue}) angezeigt werden, dann bleibt
 * der Absatz vorhanden oder über {@link #insertNoFile()} kann auch der Absatz mit entfernt werden
 * bei dem keine Einfügeoperation vorgenommen werden soll.</p>
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
public abstract class MicrosoftInsertDoc implements Extension {
    
    /*
     * Liste aller Unterstüzten Formate die eingefügt/importiert werden können.
     */
    public static enum ImportFormatType {
        
        /* Extensible HyperText Markup Language File (.xhtml). */
        XHTML ("xhtml", "application/xhtml+xml"),
        
        /* MHTML Document (.mht). */
        MHT ("mht", "application/x-mimearchive"),
        
        /* application/xml (.xml). */
        XML ("xml", "text/xml"),
        
        /* Text (.txt). */
        TEXT_PLAIN ("txt", "text/plain"),
        
        /* Wordprocessing (.docx). */
        WORD_PROCESSING_ML (
                "docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
        
        /* Office Word Macro Enabled (.docm). */
        OFFICE_WORD_MACRO_ENABLED (
                "docm",
                "application/vnd.ms-word.document.macroEnabled.12"),
        
        /* Office Word Template (.dotx). */
        OFFICE_WORD_TEMPLATE (
                "dotx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.template+xml"),
        
        /* Office Word Macro Enabled Template (.dotm). */
        OFFICE_WORD_MACRO_ENABLED_TEMPLATE (
                "dotm",
                "application/vnd.ms-word.template.macroEnabled.12"),
        
        /* Rich Text Foramt (.rtf). */
        RTF ("rtf", "text/rtf"),
        
        /* HyperText Markup Language File (.htm). */
        HTML ("html", "text/html");
        
        ////////////////////////////////////////////////////////////////////////////////////////////
        
        /**
         * Ermittelt das Import-Format anhand der Dateinamensendung im übergebenen Dateinamen/-pfad
         * unabhängig der Groß- und Kleinschreibung.
         * 
         * @param fileName      Dateiname oder Dateipfad, darf nicht {@code null} sein.
         * 
         * @return              Wurde ein mögliches Format gefunden wird {@code Optional.of(..)}
         *                      mit dem Enum-Wert zurückgegeben, ansonsten {@code Optional.empty()}.
         */
        public static Optional<ImportFormatType> byFileExtension(String fileName) {
            requireNonNull(fileName, "fileName");
            
            final String pFileName = fileName.trim().toLowerCase();
            for (ImportFormatType anyChunkFormat : values()) {
                if (pFileName.endsWith("." + anyChunkFormat.fileExtension)) {
                    return Optional.of(anyChunkFormat);
                }
            }
            
            return Optional.empty();
        }
        
        /**
         * Ermittelt das Import-Format anhand der Dateinamensendung im übergeben Dateipfad
         * unabhängig der Groß- und Kleinschreibung.
         * 
         * <p>Verhält sich identisch zu: {@link #byFileExtension(String)}.</p>
         * 
         * @param filePath      Dateipfad, darf nicht {@code null} sein.
         * 
         * @return              Format oder {@code Optional.empty()}.
         */
        public static Optional<ImportFormatType> byFileExtension(Path filePath) {
            return byFileExtension(requireNonNull(filePath).toString());
        }
        
        ////////////////////////////////////////////////////////////////////////////////////////////
        
        private final String fileExtension;
        private final String mimeType;
        
        private ImportFormatType(String fileExtension, String mimeType) {
            this.fileExtension = fileExtension;
            this.mimeType = mimeType;
        }
        
        /**
         * Rückgabe der Dateinamenserweiterung kleingeschrieben ohne Punkt.
         * 
         * @return  z.B. {@code 'docx'}
         */
        protected String getFileExtension() {
            return fileExtension;
        }
        
        /**
         * Rückgabe des MIME-Type bezogen auf das Dateiformat.
         * 
         * @return  z.B. {@code 'application/xhtml+xml'}
         */
        protected String getMimeType() {
            return mimeType;
        }
        
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private static final ExtendedValue EMPTY_INSERTION = new EmptyAltChunkExtValue();
    
    /**
     * Erzeugt einen erweiterten Wert der ein ganzes Dokument repräsentieren soll, welches in der
     * späteren Vorlage/ Dokument eingefügt werden soll.
     * 
     * <p>Der übergeben Pfad {@code docToImport} zum Dokument kann eine von {@code format}
     * abweichende Dateiendung besitzen; muss dies aber nicht. Es wird vor der Erzeugung des
     * erweiterten Wertes geprüft, ob das übergebene Dokument auch wirklich existent ist.</p>
     * 
     * @param docToImport   Pfad zum Dokument das importiert werden soll; Dateinamenserweiterung ist
     *                      nicht von Bedeutung, da die Angabe in {@code format} Vorrang besitzt.
     *                      Darf nicht {@code null} sein.
     * 
     * @param format        Format des Dokumentes, welches eingefügt werden soll. Diese Angabe
     *                      muss unbedingt korrekt sein! Angabe von {@code null} nicht zulässig.
     * 
     * @return              Erweiterter-Wert der jedoch nicht das Dokument beinhaltet, sondern nur
     *                      den Pfad zum und das Format vom Dokument. Als alternative Zeichenkette
     *                      werden Pfad und Format notfalls erzeugt. Es wird nie {@code null}
     *                      zurückgegeben.
     * 
     * @throws  UncheckedIOException 
     *          Für den Fall, das am übergebenen Dateipfad (zum Dokument) sich kein Dokument
     *          befindet.
     * 
     * @throws  NullPointerException
     *          Wenn einer oder alle übergebene Parameter {@code null} sind.
     */
    public static ExtendedValue insertFile(Path docToImport, ImportFormatType format)
            throws UncheckedIOException
    {
        requireNonNull(docToImport, "docToImport");
        requireNonNull(format, "format");
        
        return new AltChunkExtValue(docToImport, format);
    }
    
    /**
     * Vereinfachte Methode zum Erzeugen eines erweiterten Wertes, der ein ganzes Dokument
     * repräsentieren soll, welches in der späteren Vorlage/ Dokument eingefügt werden soll.
     * 
     * <p>Diese Methode ist eine Vereinfachung unter der Bedingung das anhand des Dateinamens auch
     * das korrekte Format erkannt werden kann. Siehe bezüglich dem genaueren Verhalten die Methode
     * {@link #insertFile(Path, ImportFormatType)}.</p>
     * 
     * <p>Ist die Dateinamenserweiterung unbekannt, wird eine {@link IllegalStateException}
     * geworfen. Die Exceptions der Methode {@link #insertFile(Path, ImportFormatType)}
     * können weiterhin auftreten und sich hier nicht erneut aufgeführt worden.</p>
     * 
     * @param docToImport   Pfad zum Dokument das importiert werden soll; die Dateinamenserweiterung
     *                      muss unbedingt vorhanden sein und das korrekte Format wiederspiegeln!
     *                      Darf nicht {@code null} sein.
     * 
     * @return              Erweiterter-Wert (siehe andere Methode).
     * 
     * @throws  IllegalStateException
     *          Wenn das Dateiformat (aus der Dateinamenserweiterung) unbekannt ist oder nicht
     *          konkret ermittel bar ist. Die Exception-Message ist primär intern gedacht.
     */
    public static ExtendedValue insertFile(Path docToImport) {
        requireNonNull(docToImport, "docToImport");
        
        final Optional<ImportFormatType> opDocFormat = ImportFormatType
                .byFileExtension(docToImport);
        
        if (opDocFormat.isPresent() == false) {
            throw new IllegalStateException(formatMessage(
                    "MicrosoftInsertDoc_UnknownFileFormat", docToImport));
        }
        
        return insertFile(docToImport, opDocFormat.get());
    }
    
    /**
     * Erstellt einen erweiterten Wert für einen Einsetzungsvorgang in dem KEINE Einfügeopration
     * ausgefhrt werden soll und der Absatz (mit dem Platzhalter) vollständig entfernt werden soll.
     * 
     * @return  Erweiterter-Wert ohne Dokument zum Einfügen (leere Einsetzung).
     */
    public static ExtendedValue insertNoFile() {
        return EMPTY_INSERTION;
    }
    
    /**
     * Interne Darstellung des erweiterten Wertes mit AltChunkID, gecachetem Dateiinhalt und
     * angegebenem Formattyp.
     */
    private static class AltChunkExtValue extends ExtendedValue {
        
        /* Pfad zum einfügenden Dokument. */
        private final Path importFile;
        
        /* Formattyp (unabhängig der Dateiendung). */
        private final ImportFormatType format;
        
        /* UUID wird als ChunkID verwendet um grundsätzlich Eindeutigkeit zu bekommen. */
        private final UUID chunkUID;
        
        /* Ist nach dem ersten Laden mit dem Teildokument belegt und dann non-null. */
        @Nullable
        private transient byte[] cachedDocument;

        public AltChunkExtValue(Path importFile, ImportFormatType format) {
            this.importFile = requireNonNull(importFile, "importFile");
            this.format = requireNonNull(format, "format");
            
            checkFileExists(importFile);
            
            this.chunkUID = UUID.randomUUID();
        }
        
        private void checkFileExists(Path file) {
            if (Files.exists(file) == false) {
                throw new UncheckedIOException(new IOException(createExceptionMessage()));
            }
        }
        
        private String createExceptionMessage() {
            return formatMessage("MicrosoftInsertDoc_FailedToLoad", importFile);
        }
        
        /**
         * Lädt das Dokument welches unter dem Pfad angegeben ist und cached es.
         * 
         * @return  Array aus bytes mit dem Dateiinhalt, nie {@code null}
         * 
         * @throws  UncheckedIOException
         *          Gewrappte IOException bei einem Ein-/Ausgabefehler.
         */
        protected byte[] loadImportDocument() {
            if (null != this.cachedDocument) {
                return Arrays.copyOf(cachedDocument, cachedDocument.length);
            }
            
            try {
                this.cachedDocument = Files.readAllBytes(importFile);
            } catch (IOException ioException) {
                throw new UncheckedIOException(ioException);
            }
            
            return Arrays.copyOf(cachedDocument, cachedDocument.length);
        }

        /**
         * AltChunk Relationship-ID die eindeutig über mehrere Dokumente ist.
         * 
         * <p>Der internen UUID (ChunkUID) werden die Bindestriche zuvor entfernt.</p>
         * 
         * @return  z.B. {@code 'altChunk75886d8ee2084c10b151f6b52bce9174'}, nie {@code null}.
         */
        protected String getChunkUID() {
            return ("altChunk" + chunkUID.toString()).replace("-", "");
        }
        
        /**
         * Interner Dateiname mit Bezug zur ChunkID.
         * 
         * @return  z.B. {@code 'altChunk75886d8ee2084c10b151f6b52bce9174.docx'}.
         */
        protected String getTargetName() {
            return getChunkUID() + "." + format.fileExtension;
        }

        /**
         * Pfad zur importierenden/einzufügenden Datei/Dokument.
         * 
         * @return  nie {@code null}.
         */
        protected Path getImportFile() {
            return importFile;
        }

        /**
         * Format der einzufügenden Datei/ des einzufügenden Dokumentes unabhängig von der
         * Dateinamenserweiterung.
         * 
         * @return  nie {@code null}.
         */
        protected ImportFormatType getFormat() {
            return format;
        }

        /**
         * Rückgabe des Schnittstellennamens, des Formates und der Pfad zum einzufügenden Dokument.
         * 
         * @return  z.B. {@code "ExtendedValue[XML://'C:\Users\admin\Desktop\insert.xml']"}
         */
        @Override
        public String altString() {
            return ExtendedValue.class.getSimpleName()
                    + "[" + format + "://\'" + importFile + "\']";
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public int hashCode() {
            return Objects.hash(AltChunkExtValue.class, this.importFile, this.format);
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public boolean equals(Object obj) {
            if (this == obj) {
                return true;
            }
            
            if (obj == null || getClass() != obj.getClass()) {
                return false;
            }

            final AltChunkExtValue other = (AltChunkExtValue) obj;
            
            if (!Objects.equals(this.importFile, other.importFile)) {
                return false;
            }
            
            return this.format == other.format;
        }
        
    }
    
    /**
     * Erweiterter Wert der einer nicht vorhandenen Ersetzung entspricht, dessen Absatz jedoch
     * (in dem das Platzhalter dazu angegeben wurde) mit entfernt werden soll.
     */
    private static class EmptyAltChunkExtValue extends ExtendedValue {

        @Override
        public String altString() {
            return "";
        }
        
        @Override
        public String toString() {
            return altString();
        }
        
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    /**
     * Zulässige Parent-Elemente in WordML die einen w:altChunk Element als Kind-Element beinhalten
     * dürfen ohne invalides WordML/XML zu generieren.
     * 
     * <p>Entspricht einer Teilmenge die identisch mit {@code 'w:p'} (Absatz) ist.</p>
     */
    private static final String[] ALLOWED_ALT_CHUNK_PARENTS = {
        "w:body", "w:comment", "w:docPartBody", "w:endnote", "w:footnote", "w:ftr", "w:hdr", "w:tc"
    };
    
    /** Namespace für AltChunk Elemente, Attribute und co. */
    private static final String NS_ALT_CHUNK =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";
    
    /** [word/_rels/document.xml.rels]://Relationships */
    private static final String EL_RELATIONSHIPS = "Relationships";
    
    /** [word/_rels/document.xml.rels]://Relationships[Relationship] */
    private static final String EL_RELATIONSHIP = "Relationship";
    
    ////////////////////////////////////////////////////////////////////////////////////////////////

    /**
     * Interner Konstruktor - nicht verwenden.
     */
    protected MicrosoftInsertDoc() { }
    
    protected abstract Document getWordDocument();
    protected abstract Document getRelationshipDocument();
    protected abstract Document getContentTypeDocument();

    /**
     * Implementiert das hinzufügen einer "Unter"-Datei zum eigentlichen Dokument.
     * 
     * <p>Implementierende Klasse muss dies entweder als ZIP ({@code *.docx}) oder besonders
     * kodoeirt bei Packaged-WordML Dateien ({@code *.xml}).</p>
     * 
     * @param partName  Pfad im Document-Container
     * @param data      Dokument
     */
    protected abstract void overwritePartInContainer(String partName, byte[] data);
    
    /**
     * Regristriert und fügt w:altChunk als Element an die Stelle des Absatzes in dem der
     * Platzhalter vorkommt.
     * 
     * <p>Übernimmt die ggf. notwendigen Eintragungen in {@code '[Content_Types].xml'} und in
     * {@code 'word/_rels/document.xml.rels'} sowie direkt des Elementes in
     * {@code 'word/document.xml'}.</p>
     * 
     * @param anyNode       Element oder Unterelement von {@code 'w:instrText'}.
     * 
     * @param extValue      Erweiterter Wert der dieser Erweiterung zu entsprechen hat.
     */
    public void insertAltChunkAt(Node anyNode, ExtendedValue extValue) {
        assert null != extValue : "extValue == null";
        assert isAltChunkExtValue(extValue) : "(extValue instanceof AltChunkExtValue) == false";
        
        // Prüfen auf Sonderfall einer leeren Ersetzung
        if (extValue instanceof EmptyAltChunkExtValue) {
            // Es soll kein Dokument eingefügt werden, dann nur den Absatz mit dem Platzhalter
            // entfernen und keine weitere Aktion ausführen
            
            final Node wpNode = findUpperAllowedParent(anyNode);
            
            wpNode.getParentNode().removeChild(wpNode);
            
            return;
        }
        
        // Zulässigen Node-Parent suchen der so nahe liegt wie möglich, in diesem dann w:altChunk
        // anlegen und abhängig der r:id eintragen.
        
        final AltChunkExtValue chunkExtValue = (AltChunkExtValue) extValue;
        final Node paragraphNode = findUpperAllowedParent(anyNode);
        final Node allowedParentNode = paragraphNode.getParentNode();
        
        allowedParentNode.replaceChild(
                createAltChunkElement(chunkExtValue.getChunkUID()),
                paragraphNode);
        
        
        // Im Relationships-Dokument eine neue r:id einfügen, nur wenn diese noch nicht enthalten
        // ist; eine bereits eingetragene wird nicht doppelt eingetragen.
        // Ist die ChunkID noch nicht eingetragen gewesen, wird auch das Teildokument dem
        // (ZIP) Container hinzugefügt.
        if (registerAndInsertChunkRelationship(chunkExtValue)) {
            // Auch in Content-Types mit aufführen anhand der Dateinamensendung
            registerMimeType(chunkExtValue);
        }
    }
    
    /**
     * Sucht an der Stelle des Platzhalter-Nodes, in den höherwertigen Parent-Nodes nach einem
     * passenden Elternelement für w:altChunk.
     * 
     * <pre>
     * Strukturelles Beispiel (vereinfachte XML Baumdarstellung):
     * 
     *  w:tr
     *      w:tc
     *          w:tcPr
     *      w:p
     *          w:r
     *              w:t
     *      w:p             &lt;---- AUSGABE (w:p kann mit w:altChunk getauscht werden)
     *          w:r
     *              w:fldChar
     *          w:r
     *              w:instrText         &lt;---- EINGABE
     *          w:r
     *              w:fldChar
     * </pre>
     * 
     * @param placeHolderNode   Irgendein Node vom Platzhalter-Element
     * 
     * @return                  Node der mit {@code w:altChunk} ausgetauscht werden kann.
     * 
     * @throws  IllegalStateException
     *          Wenn wirklich kein Node gefunden wurde der passend ist (unwahrscheinlich).
     */
    private Node findUpperAllowedParent(Node placeHolderNode) {
        assert null != placeHolderNode : "placeHolderNode == null";
        
        final Node allowedParent = findUpperAllowedParent0(placeHolderNode);
        
        if (null == allowedParent) {
            throw new IllegalStateException("No allowed parent node für w:altChunk found");
        }
        
        return allowedParent;
    }
    
    @Nullable @CheckReturnValue
    private Node findUpperAllowedParent0(Node node) {
        if (isValidParent(node.getParentNode())) {
            return node; // Perfekt!
        }
        
        // Dieser DOM Node ist nicht möglich, dann Parent-Node suchen und von der Position
        // n (von node) dekrementieren bis n == 0 ist um einen passenden Node zu finden.
        
        final Node parentNode = node.getParentNode();
        
        if (null == parentNode) {
            // Ende des DOM erreicht, dann wird nichts gefunden
            return null;
        }
        
        final NodeList parentChildNodes = parentNode.getChildNodes();
        
        if (parentChildNodes.getLength() == 1) {
            // Wenn nur ein Kind enthalten ist, dann ist es mit Sicherheit node, damit kann direkt
            // zum nächst höheren Parent Node gesprungen werden.
            return findUpperAllowedParent0(parentNode);
        }
        
        // Ist die Position von node in der Liste der Kinder vom Parent
        int indexOfNode = -1;
        for (int nodeIndex = 0; nodeIndex < parentChildNodes.getLength(); nodeIndex++) {
            final Node childNode = parentChildNodes.item(nodeIndex);
            
            if (childNode.isSameNode(node)) {
                indexOfNode = nodeIndex;
                break;
            }
        }
        
        // Diese Assertion _dürfte_ eigentlich nie auftreten, eigentlich...
        assert indexOfNode != -1 : "Illegal State: Child-Node is not element of Parent-Child-Nodes";
        
        // Alle Child-Nodes von indexOfNode..0 iterieren und prüfen ob die ein gültiger Parent für
        // w:altChunk sein können
        for (int nodeIndex = (indexOfNode - 1); nodeIndex > 0; nodeIndex--) {
            final Node childNode = parentChildNodes.item(nodeIndex);
            
            if (isValidParent(childNode)) {
                return childNode;
            }
        }
        
        return findUpperAllowedParent0(parentNode);
    }

    /**
     * Muss einem in {@link #ALLOWED_ALT_CHUNK_PARENTS} definierten Element-Bezeichnern entsprechen
     * und selbst vom Typ {@link Node#ELEMENT_NODE} sein.
     * 
     * @param node  Irgendein DOM Node
     * 
     * @return      nur {@code true}, wenn alle Bedingungen erfüllt sind
     */
    private boolean isValidParent(Node node) {
        if (node.getNodeType() != Node.ELEMENT_NODE) {
            return false;
        }
        
        final String nodeName = nullToEmpty(node.getNodeName());
        for (String allowedElementName : ALLOWED_ALT_CHUNK_PARENTS) {
            if (allowedElementName.equalsIgnoreCase(nodeName)) {
                return true;
            }
        }
        
        return false;
    }
    
    private String nullToEmpty(String value) {
        return null == value ? "" : value;
    }
    
    /**
     * Trägt die AltChunkID und den Speicherort in der Relationship-XML, wenn notwendig, ein.
     * 
     * @param chunkExtValue     Erweiterter Wert
     * 
     * @return                  {@code true}, wenn eine Eintragung erfolgen musste.
     */
    private boolean registerAndInsertChunkRelationship(AltChunkExtValue chunkExtValue) {
        final Document docRel = getRelationshipDocument();
        final Node nodeRelRoot = docRel.getElementsByTagName(EL_RELATIONSHIPS).item(0);
        final NodeList relChilds = nodeRelRoot.getChildNodes();
        final String chunkID = chunkExtValue.getChunkUID();
        
        for (int i = 0; i < relChilds.getLength(); i++) {
            final Node childNode = relChilds.item(i);
            
            final boolean isElementType = childNode.getNodeType() == Node.ELEMENT_NODE;
            final boolean isRelElement = EL_RELATIONSHIP.equalsIgnoreCase(childNode.getNodeName());
            
            if ((isElementType && isRelElement) == false) {
                // Wenn es kein Element und/oder kein Relationship Eintrag ist, dann können wir
                // gleich zum nächsten Child-Node springen.
                continue;
            }
            
            final NamedNodeMap attributes = childNode.getAttributes();
            final Node attrId = attributes.getNamedItem("Id");

            if (null == attrId) {
                continue;
            }

            final String idValue = attrId.getNodeValue();
            final boolean idIsAlreadyPresent = chunkID.equals(idValue);

            if (idIsAlreadyPresent) {
                // Die ID (chunkID) ist bereits eingetragen (ggf. dadurch das dieser erweiterte
                // Wert bereits woanders im Dokument eingesetzt wurde) und muss dem entsprechend
                // nicht erneut eingetragen werden und auch nicht erneut dem Dokument
                // hinzugefügt werden.
                //
                // Das Eintragen in die [Content-Types].xml kann damit auch entfallen, das dies
                // ansonsten eh passiert wäre.
                return false;
            }
        }
        
        
        // Die Chunk-ID ist noch nicht in Relationships eingetragen und auch noch nicht in der
        // ZIP Datei enthalten ...
        final Node newRelNode = createRelationshipElement(chunkExtValue);
        nodeRelRoot.appendChild(newRelNode);
        
        overwritePartInContainer(
                "word/" + chunkExtValue.getTargetName(),
                chunkExtValue.loadImportDocument());
        
        return true;
    }
    
    /**
     * Registriert den MIME-Type in der Content-Types-Datei sobald der Dateityp noch unbekannt ist.
     * 
     * @param extValue  Erweiterter Wert
     */
    private void registerMimeType(AltChunkExtValue extValue) {
        final String FILE_EXT = extValue.getFormat().getFileExtension();
        
        final Document doc = getContentTypeDocument();
        final NodeList typeNodes = doc.getElementsByTagName("Default");
        
        for (int nodeIndex = 0; nodeIndex < typeNodes.getLength(); nodeIndex++) {
            final Node node = typeNodes.item(nodeIndex);
            final NamedNodeMap attributes = node.getAttributes();
            
            final String extension = attributes.getNamedItem("Extension").getTextContent();
            if (FILE_EXT.equalsIgnoreCase(extension)) {
                // Bereits vorhanden!!
                return;
            }
        }
        
        final Node rootNode = doc.getElementsByTagName("Types").item(0);
        final Element extensionNode = doc.createElement("Default");
        
        extensionNode.setAttribute("Extension", extValue.getFormat().getFileExtension());
        extensionNode.setAttribute("ContentType", extValue.getFormat().getMimeType());
        
        rootNode.insertBefore(extensionNode, rootNode.getFirstChild());
    }
    
    private Node createAltChunkElement(String chunkID) {
        final Document doc = getWordDocument();
        final Element elementAltChunk = doc.createElement("w:altChunk");
        
        elementAltChunk.setAttribute("r:id", chunkID);
        
        return elementAltChunk;
    }
    
    private Node createRelationshipElement(AltChunkExtValue extValue) {
        final Document doc = getRelationshipDocument();
        
        final Element elementRel = doc.createElement(EL_RELATIONSHIP);
        elementRel.setAttribute("Id", extValue.getChunkUID());
        elementRel.setAttribute("TargetMode", "Internal");
        elementRel.setAttribute("Type", NS_ALT_CHUNK);
        elementRel.setAttribute("Target", extValue.getTargetName());
        
        return elementRel;
    }

    /**
     * Prüft ob der übergebene erweiterte Wert vom Typ dieser Erweiterung stammt.
     * 
     * @param extValue  Irgendein erweiterter Wert (nicht {@code null})
     * 
     * @return          {@code true}, wenn die Wert zu dieser Erweiterung zugehörig ist.
     */
    public boolean isAltChunkExtValue(ExtendedValue extValue) {
        if (null == extValue) {
            return false;
        }
        
        return extValue instanceof AltChunkExtValue
                || extValue instanceof EmptyAltChunkExtValue;
    }
    
}
