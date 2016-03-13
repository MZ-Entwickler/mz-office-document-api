/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2016,   Moritz Riebe     (moritz.riebe@mz-solutions.de)
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
package com.mz.solutions.office;

import com.mz.solutions.office.OfficeDocumentException.DocumentPlaceholderMissingException;
import com.mz.solutions.office.OfficeDocumentException.NoDataForDocumentGenerationException;
import com.mz.solutions.office.extension.Extension;
import com.mz.solutions.office.extension.MicrosoftCustomXml;
import com.mz.solutions.office.model.DataMap;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataTable;
import com.mz.solutions.office.model.DataTableMap;
import com.mz.solutions.office.model.DataTableRow;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.model.DataValueMap;
import com.mz.solutions.office.model.ValueOptions;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.StringTokenizer;
import javax.annotation.Nullable;
import static mz.solutions.office.resources.MessageResources.formatMessage;
import mz.solutions.office.resources.MicrosoftDocumentKeys;
import static mz.solutions.office.resources.MicrosoftDocumentKeys.UNKNOWN_FORMATTING_CHAR;
import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

final class MicrosoftDocument extends AbstractOfficeXmlDocument {
    
    private static final String ZIP_DOC_DOCUMENT = "word/document.xml";
    private static final String ZIP_DOC_STYLES = "word/styles.xml";
    
    /**
     * Speichert die Referenz wenn die Custom XML Erweiterung verwendet wurde;
     * wurde die Erweiterung nicht verwendet, bleibt das Feld mit null belegt.
     */
    @Nullable
    private MicrosoftCustomXml extCustomXml;

    public MicrosoftDocument(OfficeDocumentFactory factory, Path document) {
        super(factory, document, ZIP_DOC_DOCUMENT, ZIP_DOC_STYLES);
    }

    @Override
    protected String getImplementedOfficeName() {
        return "Microsoft Office 2007 - 2016";
    }

    @Override
    public <T extends Extension> Optional<T> extension(Class<T> extType) {
        // Siehe überschriebene Methode warum super aufgerufen wird
        final Optional<T> superExtension = super.extension(extType);
        if (superExtension.isPresent()) {
            return superExtension;
        }
        
        if (extType == MicrosoftCustomXml.class) {
            if (null == extCustomXml) {
                extCustomXml = new ZippedCustomXmlExtension();
            }
            return Optional.of((T) extCustomXml);
        }
        
        return Optional.empty();
    }
    
    /**
     * Überprüft Notwendigkeit ob zwischen DataPages ein Seitenumbruch
     * eingefügt werden soll.
     * 
     * @return      bei {@code true} einfügen
     */
    private boolean needToInsertPageBreak() {
        Boolean doInsertPageBreaks = myOfficeFactory
                .getProperty(MicrosoftProperty.INS_HARD_PAGE_BREAKS);
        
        return Boolean.TRUE.equals(doInsertPageBreaks);
    }

    @Override
    protected ZIPDocumentFile createAndFillDocument(
            final Iterator<DataPage> dataPages) {
        
        final ZIPDocumentFile newFile = sourceDocumentFile.cloneDocument();
        
        fillDocuments(dataPages, newFile);
        
        return newFile;
    }
    
    private void fillDocuments(
            final Iterator<DataPage> dataPages,
            final ZIPDocumentFile outputDocument) {
        
        final Document newContent = (Document) sourceContent.cloneNode(true);
        final Document newStyles = (Document) sourceStyles.cloneNode(true);
        
        normalizeInstrTextFields(newContent);
        
        final Node wordBody = findDocumentBody(newContent);
        final Node newWordBody = wordBody.cloneNode(true);
        
        while (newWordBody.hasChildNodes()) {
            newWordBody.removeChild(newWordBody.getFirstChild());
        }
        
        boolean firstPage = true;
        boolean missingData = true;
        
        while (dataPages.hasNext()) {
            final DataPage pageData = dataPages.next();
            final Node newPageBody = wordBody.cloneNode(true);
            
            replaceAllFields(newPageBody, pageData);
            
            while (newPageBody.hasChildNodes()) {
                newWordBody.appendChild(
                        newPageBody.removeChild(newPageBody.getFirstChild())
                );
            }
            
            if (firstPage == false && needToInsertPageBreak()) {
                final Node wordBreak = newContent.createElement("w:br");
                final NamedNodeMap attributes = wordBreak.getAttributes();
                
                final Attr wordType = newContent.createAttribute("w:type");
                wordType.setValue("page");
                
                attributes.setNamedItem(wordType);
                
                newWordBody.appendChild(wordBreak);
            }
            
            missingData = false;
            firstPage = false;
        }
        
        if (missingData) {
            if (ignoreMissingDataPages()) {
                return; // Leise beenden
            }
            
            throw new NoDataForDocumentGenerationException(
                    formatMessage(MicrosoftDocumentKeys.NO_DATA));
        }
        
        final Node wordBodyParent = wordBody.getParentNode();
        wordBodyParent.insertBefore(newWordBody, wordBody);
        wordBodyParent.removeChild(wordBody);
        
        
        removeAllBookmarkTags(newContent);
        
        overwrite(outputDocument, ZIP_DOC_DOCUMENT, newContent);
        overwrite(outputDocument, ZIP_DOC_STYLES, newStyles);
        
        //// EXTENSIONS DURCHLAUFEN LASSEN
        
        if (null != extCustomXml) {
            assert extCustomXml instanceof ZippedCustomXmlExtension;
            ((ZippedCustomXmlExtension)extCustomXml)
                    .applyCustomXmlData(outputDocument);
        }
    }
    
    private void normalizeInstrTextFields(Node rootNode) {
        // Alle Field-Chars suchen (somit Begin & End)
        // Filtern nach den Field-Chars die NUR Beginn-Zeichen markieren
        streamElements(rootNode)
                .filter(n -> isElement(n, "w:fldChar"))
                .filter(n -> hasAttribute(n, "w:fldCharType"))
                .filter(n -> getAttribute(n, "w:fldCharType").equals("begin"))
                .forEach(this::collapseInstrTextField);
    }
    
    private void collapseInstrTextField(Node fieldCharNode) {
        //  <w:p>
        //      <w:pPr>...</w:pPr>
        //      <w:r>
        //          <w:rPr>...</w:rPr>
        //          <w:fldChar w:fldCharType="begin"/>       -> := fieldCharNode
        //      </w:r>
        //      <w:r>
        //          <w:rPr>...</w:rPr>
        //          <w:instrText>PLACEHOLDER</w:instrText>
        //      </w:r>
        //      ... weitere instrText
        //      <w:r>
        //          <w:rPr>...</w:rPr>
        //          <w:fldChar fldCharType="separate"/>
        //      </w:r>
        //      ... aktueller Anzeigewert                   -> muss weg
        //      <w:r>
        //          <w:rPr>...</w:rPr>
        //          <w:fldChar fldCharType="end"/>          -> bis hier
        //      </w:r>
        
        final Node wordRun = fieldCharNode.getParentNode();
        final Node wordP = wordRun.getParentNode();
        
        final List<Node> wordPChilds = toFlatNodeList(wordP.getChildNodes());
        final int wordPChildLen = wordPChilds.size();
        
        // Jetzt werden NUR w:r rausgesucht die auch uns betreffen und zwischen
        // dem Beginn und Endzeichen liegen -> der Rest interessiert uns nicht
        final List<Node> ourChilds = new LinkedList<>();
        
        // Suche Start-Index für UNSEREN wordRun (w:r) mit dem Startzeichen
        int startIndex = -1;
        for (int i = 0; i < wordPChildLen; i++) {
            if (wordPChilds.get(i) == wordRun /* unserer */) {
                startIndex = i;
                break;
            }
        }
        
        // Alle w:r zusammensammeln bis fldChar="end" kommt (jenes inklusiv)
        for (int i = startIndex; i < wordPChildLen; i++) {
            final Node wR = wordPChilds.get(i);
            final Node fldChar = findNodeByName(wR, "w:fldChar");
            
            ourChilds.add(wR);
            
            // Ist in wR ein fldChar drinne? dann ist fldChar != null
            if (null == fldChar) {
                continue; // anscheind nicht, nächstes w:r
            }
            
            // wenn doch, dann muss das fldCharType="end" haben!
            if (getAttribute(fldChar, "w:fldCharType").equals("end")) {
                // Treffer! Hier endet UNSER Field-Code!
                break;
            }
        }
        
        // ZWISCHENSTAND
        // in ourChilds sind alle <w:r> Elemente drinne die zu unserem
        // Platzhalter gehören. Davon gibt es mindestens ein <w:instrText>
        // und mindestens zwei <w:fldChar>
        
        // Fall 1:  es gibt nur EIN instrText -> dann muss hier nichts
        //          mehr zusammengesetzt werden, Platzhalter ist normalisiert.
        //          Ausgenommen von mehr als 2 fldChar's, dann muss der ehemals
        //          angezeigte Text rausgenommen werden
        
        final List<Node> instrTextNodes = new LinkedList<>();
        final List<Node> fldCharNodes = new LinkedList<>();
        
        for (Node anyWordRun : ourChilds) { // Suchen alle w:instrText zusammen!
            streamElements(anyWordRun)
                    .filter(n -> isElement(n, "w:instrText"))
                    .forEach(instrTextNodes::add);
            
            streamElements(anyWordRun)
                    .filter(n -> isElement(n, "w:fldChar"))
                    .forEach(fldCharNodes::add);
        }
        
        final boolean isAlreadyNormalized =
                instrTextNodes.size() == 1   // 1x Platzhalterinhalt
                && fldCharNodes.size() == 2; // 2x Klammern des Feldbefehls
        
        if (isAlreadyNormalized) {
            return; // -> nichts mehr für uns zu tun, Fall 1 trifft zu
        }
        
        // Fall 2:  Es gibt mehrere w:instrText, deren Inhalt setzen wir
        //          zusammen und verpacken es im ersten instrText
        
        final Node firstInstrText = instrTextNodes.get(0);
        final StringBuilder fullTextBuilder = new StringBuilder();
        
        for (Node anyInstrText : instrTextNodes) {
            fullTextBuilder.append(getTextContent(anyInstrText));
        }
        
        firstInstrText.getChildNodes().item(0) // text von instrText
                .setNodeValue(fullTextBuilder.toString());
        
        // AUFRÄUMEN
        // Alle w:r entfernen die:
        //  a) kein fldCharType="begin" enthalten
        //  b) kein fldCharType="end" enthalten
        //  c) ungleich firstInstrText enthalten
        
        // i = 0 -> field begin
        // i = 1 -> erster instrText
        // i - 1 -> field end
        
        for (int i = 2; i < ourChilds.size() - 1; i++) {
            wordP.removeChild(ourChilds.get(i));
        }
    }
    
    private void replaceAllFields(Node rootNode, DataMap values) {
        replaceFields(rootNode, values);
        
        for (Node tableNode : walkTablesFlat(rootNode)) {
            final Optional<String> tableName = findKnownTableBookmarkName(
                    tableNode, values);
            
            final boolean isUnkownTable = tableName.isPresent() == false;
            
            if (isUnkownTable) {
                // Unbekannte Tabellen werden mit selben Daten befüllt
                replaceFields(tableNode, values);
            } else {
                // Bekannte Tabellen werden korrekt mit den Datenzeilen befüllt
                final Optional<DataTable> tableValues =
                        values.getTableByName(tableName.get());
                
                replaceTable(tableNode, tableValues.get());
            }
        }
    }
    
    private void replaceFields(Node rootNode, DataValueMap values) {
        for (Node instrTextNode : walkInstrTextsButNoTables(rootNode)) {
            replaceField(instrTextNode, values);
        }
    }
    
    private void replaceField(Node instrTextNode, DataValueMap values) {
        // Nach Platzhalterbezeichner suchen
        final NodeList childs = instrTextNode.getChildNodes();
        String keyName = null;
        
        for (int i = 0; i < childs.getLength(); i++) {
            final Node keyNode = childs.item(i);
            if (keyNode.getNodeType() == Node.TEXT_NODE) {
                keyName = keyNode.getNodeValue().trim().toUpperCase();
            }
        }
        
        if (null == keyName) {
            throw new IllegalStateException("keyName == null ?!?!?!");
        }
        
        // Field Code mit Platzhalter?
        if (keyName.contains("DOCVARIABLE")) {
            keyName = keyName.replace("DOCVARIABLE", "")
                    .replace('\"', ' ').trim();
        }
        
        if (keyName.contains("MERGEFIELD")) {
            keyName = keyName.replace("MERGEFIELD", "")
                    .replace('\"', ' ').trim();
            
            final StringTokenizer t = new StringTokenizer(keyName, " \\*");
            if (t.hasMoreTokens() == false) {
                throw new DocumentPlaceholderMissingException(formatMessage(
                        MicrosoftDocumentKeys.INVALID_MERGE_FIELD));
            }
            
            keyName = t.nextToken();
        }
        
        // Wert zum Platzhalter suchen
        final Optional<DataValue> value = values.getValueByKey(keyName);
        if (value.isPresent() == false) {
            if (ignoreMissingValues()) {
                return; // Soll ignoriert werden
            }

            // Vor der Exception, mmuss geprüft werden ob es sich um einen
            // potentiellen Feldbefehl handelt, oder einen Platzhalter
            
            // Zeichen die auf einen Word Feldbefehl hindeuten
            final char[] creepyChars = " \"*\'+-!#\\".toCharArray();
            final String pKeyName = keyName;
            
            boolean containsCreepyChar = false;
            for (char ch : creepyChars) {
                if (pKeyName.contains(Character.toString(ch))) {
                    containsCreepyChar = true;
                    break;
                }
            }
            
            if (containsCreepyChar == false) {
                // Sieht schlecht aus, sehr wahrscheinlich Platzhalter
                throw new DocumentPlaceholderMissingException(formatMessage(
                        MicrosoftDocumentKeys.UNKNOWN_PLACE_HOLDER,
                                /* {0} */ pKeyName));
            }
            
            return; // Joar anderweitiger Feldbefehl
        }
        
        // w:instrText ersetzen durch w:t oder längerer Formatierungskette
        final Node wordRun = instrTextNode.getParentNode();
        final List<Node> formattedNodes = createFormattedNodes(
                instrTextNode, value.get());
        
        for (Node formattedNode : formattedNodes) {
            wordRun.insertBefore(formattedNode, instrTextNode);
        }
        
        // w:instrText tauschen mit formattierten Nodes
        wordRun.removeChild(instrTextNode);
        
        // umgebende w:fldChar's entfernen
        // Außerhalb der w:r Tags (in denen w:instrText war), gibt es einen
        // (meinst) w:p Tag -> dieser enthält die w:fldChar's
        // >> Einfachste Lösung: Tags davor & danach entfernen
        final Node wordP = wordRun.getParentNode();
        final NodeList wordPChilds = wordP.getChildNodes();
        
        // Index von wordP finden und wordP[i-1] und wordP[i+1] entfernen
        final int wordPChildLen = wordPChilds.getLength();
        
        int wordRunIndex = -1;
        for (int i = 0; i < wordPChildLen; i++) {
            if (wordPChilds.item(i) == wordRun) {
                wordRunIndex = i;
                break;
            }
        }
        
        final int wordRunBefore = wordRunIndex - 1;
        final int wordRunAfter = wordRunIndex + 1;
        
        if (wordRunBefore >= 0 && wordRunBefore < wordPChildLen) {
            wordP.removeChild(wordPChilds.item(wordRunBefore));
        }
        
        if (wordRunAfter >= 0 && wordRunAfter < wordPChildLen) {
            wordP.removeChild(wordPChilds.item(wordRunAfter - 1));
        }
    }
    
    private List<Node> createFormattedNodes(
            Node instrTextNode, DataValue value) {
        
        final Document document = instrTextNode.getOwnerDocument();
        final Set<ValueOptions> options = value.getValueOptions();
        
        final boolean isSimpleText =
                options.contains(ValueOptions.KEEP_LINEBREAK) == false
                && options.contains(ValueOptions.KEEP_TABULATOR) == false;
        
        final String textContent = value.getValue();
        
        if (isSimpleText) {
            final Node wordText = createWordTextNode(document, textContent);
            
            return Collections.singletonList(wordText);
        }
        
        // Doch etwas komplizierter Formatiert =D
        
        final List<Node> formattedNodes = new LinkedList<>();
        final StringBuilder textBuffer = new StringBuilder();
        
        for (char ch : textContent.toCharArray()) {
            if (ch == '\r') {
                continue; // CR wird ignoriert
            }
            
            if (ch != '\n' && ch != '\t') {
                textBuffer.append(ch); // da kein Sonderzeichen
                continue;
            }
            
            // Sonderzeichen
            final boolean needTextNode = textBuffer.length() > 0;
            
            if (needTextNode) {
                final Node textPartNode = createWordTextNode(
                        document, textBuffer.toString());
                
                formattedNodes.add(textPartNode);
                
                textBuffer.setLength(0);
            }
            
            switch (ch) {
                case '\n': formattedNodes.add(document.createElement("w:br"));
                    break;
                    
                case '\t': formattedNodes.add(document.createElement("w:tab"));
                    break;
                    
                default:
                    throw new IllegalStateException(
                            formatMessage(UNKNOWN_FORMATTING_CHAR,
                                    /* {0} */ Integer.toString((int) ch)));
            }
        }
        
        if (textBuffer.length() > 0) {
            final Node wordTextNode = createWordTextNode(
                    document, textBuffer.toString());
            
            formattedNodes.add(wordTextNode);
        }
        
        return formattedNodes;
    }
    
    private Node createWordTextNode(Document document, String textContent) {
        final Node wordText = document.createElement("w:t");
        final Node textNode = document.createTextNode(textContent);
        
        wordText.appendChild(textNode);
        
        return wordText;
    }
    
    private void removeAllBookmarkTags(Node rootNode) {
        for (Node bookmarkNode : walkAllBookmarkTags(rootNode)) {
            final Node parent = bookmarkNode.getParentNode();
            
            parent.removeChild(bookmarkNode);
        }
    }
    
    private Optional<String> findKnownTableBookmarkName(
            final Node tableNode, final DataTableMap tableMap) {
        
        return walkBookmarksInTable(tableNode).asList().stream()
                .map(n -> getAttribute(n, "w:name"))
                .filter(name -> name.trim().isEmpty() == false)
                .filter(name -> name.length() > 2)
                .filter(name -> tableMap.getTableByName(name).isPresent())
                .findFirst();
    }
    
    private void replaceTable(Node tableNode, DataTable tableValues) {
        final List<Node> tableRowNodes = walkTableRows(tableNode).asList();
        
        final int tableRowCount = tableRowNodes.size();
        final int dataRowIndex = tableRowCount == 1 ? 0 : 1;
        
        // DataValue's in allen Zeilen ersetzen die NICHT die Datenzeile sind
        for (int i = 0; i < tableRowCount; i++) {
            if (dataRowIndex == i /* current */) {
                continue; // Datenzeile ignorieren
            }
            
            replaceFields(tableRowNodes.get(i), tableValues);
        }
        
        // Über alle DataRow's iterieren und in Tabelle einfügen
        final Node tableDataRowNode = tableRowNodes.get(dataRowIndex);
        
        boolean nothingInserted = true;
        for (DataTableRow dataRow : tableValues) {
            final Node newTableRow = tableDataRowNode.cloneNode(true);
            
            replaceAllFields(newTableRow, dataRow);
            
            tableNode.insertBefore(newTableRow, tableDataRowNode);
            
            nothingInserted = false;
        }
        
        // Eigentliche Datenzeile mit Platzhaltern entfernen
        tableNode.removeChild(tableDataRowNode);
        
        final boolean wasOnlyOneRow = tableRowCount == 1;
        
        if (wasOnlyOneRow && nothingInserted) {
            // Dann gibts keine Zeilen mehr -> Tabelle muss ganz weg
            
            final Node tableParent = tableNode.getParentNode();
            tableParent.removeChild(tableNode);
        }
    }
    
    ////////////////////////////////////////////////////////////////////////////
    // Einfache XML Routinen
    ////////////////////////////////////////////////////////////////////////////

    private Node findDocumentBody(Node rootNode) {
        return findNodeByName(rootNode, "w:body");
    }
    
    private GenericNodeIterator walkFieldChars(Node rootNode) {
        return new GenericNodeIterator(rootNode, "w:fldChar");
    }

    private GenericNodeIterator walkFieldInstrTexts(Node rootNode) {
        return new GenericNodeIterator(rootNode, "w:instrText")
                .noRecursionByElements("w:fldChar");
    }
    
    private GenericNodeIterator walkInstrTextsButNoTables(Node rootNode) {
        return new GenericNodeIterator(rootNode, "w:instrText")
                .noRecursionByElements("w:tbl")
                .doSkipFirstStopElement();
    }
    
    private GenericNodeIterator walkTablesFlat(Node rootNode) {
        return new GenericNodeIterator(rootNode, "w:tbl")
                .noRecursionByElements("w:tbl")
                .doSkipFirstStopElement();
    }
    
    private GenericNodeIterator walkTableRows(Node tableNode) {
        return new GenericNodeIterator(tableNode, "w:tr")
                .noRecursionByElements("w:tbl")
                .doSkipFirstStopElement();
    }
    
    private Iterable<Node> walkAllBookmarkTags(Node rootNode) {
        final List<Node> bookmarkStartNodes = new GenericNodeIterator(
                rootNode, "w:bookmarkStart").asList();
        
        final List<Node> bookmarkEndNodes = new GenericNodeIterator(
                rootNode, "w:bookmarkEnd").asList();
        
        final int size = bookmarkStartNodes.size() + bookmarkEndNodes.size();
        final List<Node> allBookmarkTags = new ArrayList<>(size);
        
        allBookmarkTags.addAll(bookmarkStartNodes);
        allBookmarkTags.addAll(bookmarkEndNodes);
        
        return allBookmarkTags;
    }
    
    private GenericNodeIterator walkBookmarksInTable(Node tableNode) {
        return new GenericNodeIterator(tableNode, "w:bookmarkStart")
                .noRecursionByElements("w:tbl")
                .doSkipFirstStopElement();
    }
    
    
    ////////////////////////////////////////////////////////////////////////////
    // IMPLEMENTIERUNG DER MICROSOFT CUSTOM XML ERWEITERUNG
    ////////////////////////////////////////////////////////////////////////////
    
    private class ZippedCustomXmlExtension extends CustomXmlExtensionBase {
        
        private final Map<String, byte[]> modifiedCustXmlParts;
        
        private final String[] partNames;
        private final boolean hasNoXmlParts;
        
        public ZippedCustomXmlExtension() {
            this.partNames = detectPartNames();
            this.hasNoXmlParts = (partNames.length == 0);
            
            if (hasNoXmlParts) {
                this.modifiedCustXmlParts = Collections.EMPTY_MAP;
            } else {
                this.modifiedCustXmlParts = new HashMap<>(partNames.length);
            }
        }
        
        private String[] detectPartNames() {
            return sourceDocumentFile
                    .findItemsStartingWith("customXml/item").stream()
                    .filter(name -> name.contains("itemProps") == false)
                    .filter(name -> name.endsWith(".xml"))
                    .sorted()
                    .toArray(String[]::new);
        }
        
        @Override
        public int countParts() {
            return partNames.length;
        }

        @Override
        protected void checkPartIndex(int index) {
            if (hasNoXmlParts) {
                throw new IllegalStateException("No Custom XML parts there");
            }
            
            if (index < 0 || index >= partNames.length) {
                throw new IndexOutOfBoundsException("index [0.."
                        + partNames.length + "]");
            }
        }

        @Override
        public byte[] partAsBytesAt(int index) {
            checkPartIndex(index);
            
            final String itemName = partNames[index];
            final boolean wasOverwritten = modifiedCustXmlParts
                    .containsKey(itemName);
            
            if (wasOverwritten) {
                // Wenn bereits durch den Nutzer verändert, dann gib auch direkt
                // die Änderung zurück
                final byte[] orignData = modifiedCustXmlParts.get(itemName);
                return Arrays.copyOf(orignData, orignData.length);
            }
            
            // Das byte[]-Array aus der ZIP muss nicht geklont werden, da die
            // Rückgabe der ZIP bereits klont
            return sourceDocumentFile.read(itemName);
        }

        @Override
        public void overwritePartAt(int index, byte[] value) {
            checkPartIndex(index);
            Objects.requireNonNull(value, "value");
            
            final String itemName = partNames[index];
            modifiedCustXmlParts.put(itemName, value);
        }
        
        protected void applyCustomXmlData(ZIPDocumentFile destination) {
            assert null != destination : "destination";
            
            if (hasNoXmlParts || modifiedCustXmlParts.isEmpty()) {
                return; // Fertig :-D
            }
            
            for (String itemName : modifiedCustXmlParts.keySet()) {
                final byte[] newXmlData = modifiedCustXmlParts.get(itemName);
                destination.overwrite(itemName, newXmlData);
            }
        }
        
    }
    
    /**
     * Statische Unterklasse die die mehrfachen Methoden zum Lesen
     * und Setzen der Custom XML Parts übernimmt.
     */
    private abstract class CustomXmlExtensionBase extends MicrosoftCustomXml {
        
        protected abstract void checkPartIndex(int index);
        
        @Override
        public String partAsStringAt(int index) {
            checkPartIndex(index);
            
            return new String(partAsBytesAt(index), StandardCharsets.UTF_8);
        }
        
        @Override
        public Document partAsDocumentAt(int index) {
            checkPartIndex(index);
            
            return bytesToXml(partAsBytesAt(index));
        }
        
        @Override
        public void overwritePartAt(int index, String value) {
            checkPartIndex(index);
            Objects.requireNonNull(value, "value");
            
            overwritePartAt(index, value.getBytes(StandardCharsets.UTF_8));
        }
        
        @Override
        public void overwritePartAt(int index, Node value) {
            checkPartIndex(index);
            Objects.requireNonNull(value, "value");
            
            overwritePartAt(index, xmlToBytes(value));
        }
        
        @Override
        public String toString() {
            return MicrosoftCustomXml.class.getSimpleName() + "["
                    + countParts() + " Parts]";
        }
        
    }

}
