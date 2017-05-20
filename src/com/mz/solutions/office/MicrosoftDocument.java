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
package com.mz.solutions.office;

import com.mz.solutions.office.OfficeDocumentException.DocumentPlaceholderMissingException;
import com.mz.solutions.office.OfficeDocumentException.NoDataForDocumentGenerationException;
import com.mz.solutions.office.extension.ExtendedValue;
import com.mz.solutions.office.extension.Extension;
import com.mz.solutions.office.extension.MicrosoftCustomXml;
import com.mz.solutions.office.extension.MicrosoftInsertDoc;
import com.mz.solutions.office.model.DataMap;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataTable;
import com.mz.solutions.office.model.DataTableMap;
import com.mz.solutions.office.model.DataTableRow;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.model.DataValueMap;
import com.mz.solutions.office.model.ValueOptions;
import com.mz.solutions.office.model.images.ExternalImageResource;
import com.mz.solutions.office.model.images.ImageResource;
import com.mz.solutions.office.model.images.ImageResourceType;
import com.mz.solutions.office.model.images.ImageValue;
import com.mz.solutions.office.model.images.LocalImageResource;
import com.mz.solutions.office.model.interceptor.InterceptionContext;
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
import java.util.UUID;
import javax.annotation.Nullable;
import static mz.solutions.office.resources.MessageResources.formatMessage;
import mz.solutions.office.resources.MicrosoftDocumentKeys;
import static mz.solutions.office.resources.MicrosoftDocumentKeys.UNKNOWN_FORMATTING_CHAR;
import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

final class MicrosoftDocument extends AbstractOfficeXmlDocument {
    
    private static final String ZIP_DOC_DOCUMENT = "word/document.xml";
    private static final String ZIP_DOC_STYLES = "word/styles.xml";
    private static final String ZIP_REL_RELS = "word/_rels/document.xml.rels";
    
    private static final String ZIP_CONTENT_TYPES = "[Content_Types].xml";
    
    /**
     * Speichert die Referenz wenn die Custom XML Erweiterung verwendet wurde;
     * wurde die Erweiterung nicht verwendet, bleibt das Feld mit null belegt.
     */
    @Nullable
    private MicrosoftCustomXml extCustomXml;
    
    /** Speichert die Referenz zur altChunk-Erweiterung, sobald diese verwendet wurde. */
    private InnerAltChunkExtension extAltChunk = new InnerAltChunkExtension();

    /** Interceptor-Context für Lazy-Callbacks, wird bei jedem Platzhaler neu initialisiert. */
    private MyInterceptionContext interceptionContext = new MyInterceptionContext();
    
    private final Document sourceRelationships;
    private final Document sourceContentTypes;
    
    /** Zählt die Anzahl der eingefügten Bilder; muss vor Ersetzungsvorgang zurückgesetzt werden. */
    private volatile int imageCounter = 0;

    public MicrosoftDocument(OfficeDocumentFactory factory, Path document) {
        super(factory, document, ZIP_DOC_DOCUMENT, ZIP_DOC_STYLES);
        
        this.sourceRelationships = loadFileAsXml(ZIP_REL_RELS);
        this.sourceContentTypes = loadFileAsXml(ZIP_CONTENT_TYPES);
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
        
        if (extType == MicrosoftInsertDoc.class) {
            return Optional.of((T) extAltChunk);
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
    
    /**
     * Prüft die Einstellung ob ein Drawing-Element einem VShape Element gegenüber
     * bevorzugt werden soll.
     * 
     * @return  {@code true}, dann nutz {@code w:drawing} soweit wie möglich.
     */
    private boolean useDrawingElementOverVShape() {
        final Boolean drawingOverVShape = myOfficeFactory.getProperty(
                MicrosoftProperty.USE_DRAWING_OVER_VML);
        
        return Boolean.TRUE.equals(drawingOverVShape);
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
        
        this.imageCounter = 16_000;
        
        final Document newContent = (Document) sourceContent.cloneNode(true);
        final Document newStyles = (Document) sourceStyles.cloneNode(true);
        
        final Document newRelationships = (Document) sourceRelationships.cloneNode(true);
        final Document newContentTypes = (Document) sourceContentTypes.cloneNode(true);
        
        extAltChunk.setCurrentZipFile(outputDocument)
                .setRelationshipDocument(newRelationships)
                .setWordDocument(newContent)
                .setContentTypesDocument(newContentTypes);
        
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
            
            // Zeilenumbruch nur Einfügen, wenn es sich NICHT um die
            // Seite handelt
            if (firstPage == false && needToInsertPageBreak()) {
                final Node wordBreak = newContent.createElement("w:br");
                final NamedNodeMap attributes = wordBreak.getAttributes();
                
                final Attr wordType = newContent.createAttribute("w:type");
                wordType.setValue("page");
                
                attributes.setNamedItem(wordType);
                
                newWordBody.appendChild(wordBreak);
            }
            
            // Ersetzen und alle Nodes in dem fortlaufenden neuen Word-Inhalt
            // übernehmen. #removeChild(..) gibt den entfernten Node zurück.
            
            replaceAllFields(newPageBody, pageData);
            
            while (newPageBody.hasChildNodes()) {
                newWordBody.appendChild(
                        newPageBody.removeChild(newPageBody.getFirstChild())
                );
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
        
        overwrite(outputDocument, ZIP_REL_RELS, newRelationships);
        overwrite(outputDocument, ZIP_CONTENT_TYPES, newContentTypes);
        
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
        // Wichtig: erst die Bilder ersetzen! Ansonsten werden Textplatzhalter ggf. mit einem Bild
        //          ersetz, und nachfolgenden wird wieder versucht die eben eingesetzt Bilder dann
        //          anhand deren Eigenschaften erneut zu ersetzen (könnte in einer Endlosschleife
        //          ungünstig enden)
        
        for (Node wDrawingNode : walkDrawingButNoTables(rootNode)) {
            replaceDrawing(wDrawingNode, values);
        }
        
        for (Node wPictNode : walkPictureButNoTables(rootNode)) {
            replacePicture(wPictNode, values);
        }
        
        // Feld-Befehle (Textplatzhalter)
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
        
        // Beim Ermitteln des Inhaltes, auf Interceptor prüfen und entsprechend zuvor den Context
        // dazu einrichten.
        interceptionContext.init(keyName, values);
        
        // Sonderfall: Ist der zurückgegebene DataValue ein erweiterter Wert, dann muss mindestens
        // geprüft werden ob diese Art und bekannt ist.
        final DataValue dataValue = handleInterception(value.get(), interceptionContext);
        if (dataValue.isExtendedValue()) {
            final ExtendedValue extValue = dataValue.extendedValue();
            
            if (extAltChunk.isAltChunkExtValue(extValue)) {
                // Platzhalter durch w:altChunk ersetzen
                subReplaceFieldWithAltChunk(instrTextNode, extValue);
                return;
            }
            
            if (null != extValue && extValue instanceof ImageValue) {
                subReplaceFieldWithImage(instrTextNode, extValue);
                return;
            }
        }
        
        // w:instrText ersetzen durch w:t oder längerer Formatierungskette
        final Node wordRun = instrTextNode.getParentNode();
        final List<Node> formattedNodes = createFormattedNodes(
                instrTextNode, dataValue);
        
        for (Node formattedNode : formattedNodes) {
            wordRun.insertBefore(formattedNode, instrTextNode);
        }
        
        // w:instrText tauschen mit formattierten Nodes
        wordRun.removeChild(instrTextNode);
        
        removeEnclosingFieldChars(wordRun);
    }
    
    private void removeEnclosingFieldChars(Node wordRun) {
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
    
    private void subReplaceFieldWithAltChunk(Node instrTextNode, ExtendedValue extValue) {
        // - instrTextNode -> Platzhalter vollständig (samt ABSATZ!) entfernen
        // - Weitere Behandlung an die Erweiterungsimplementierung geben
        final Node fieldParent = instrTextNode.getParentNode();
        final Node contentParent = fieldParent.getParentNode();
        
        extAltChunk.insertAltChunkAt(instrTextNode, extValue);
        
        // Jetzt erst entfernen
        //contentParent.removeChild(fieldParent);
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
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // HANDLING VON BILDERN IN WORD-DOKUMENTEN
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private void subReplaceFieldWithImage(Node instrTextNode, ExtendedValue extValue) {
        final ImageValue imageValue = (ImageValue) extValue;
        
        final Object[] registrationResult = registerImageResource(imageValue);
        final String imgRelId = (String) registrationResult[0];
        final boolean usedResourceKeepsExternal = (boolean) registrationResult[1];
        
        final Document document = instrTextNode.getOwnerDocument();
        final Element drawingOrPictureElement;
        
        if (useDrawingElementOverVShape()) {
            drawingOrPictureElement = createDrawingElement(document);
            
            setupDrawingElement(
                    drawingOrPictureElement, imageValue,
                    imgRelId, usedResourceKeepsExternal);
        } else {
            drawingOrPictureElement = createPictureElement(document);
            setupPictureElement(
                    drawingOrPictureElement, imageValue, imgRelId);
        }
        
        final Node wordRun = instrTextNode.getParentNode();
        wordRun.replaceChild(drawingOrPictureElement, instrTextNode);
        
        // noch die w:fldChar's finden und entfernen (ja die sind immernoch da -.-)
        removeEnclosingFieldChars(wordRun);
    }
    
    private void replaceDrawing(Node wDrawingNode, DataValueMap values) {
        // Grundvoraussetzungen damit ein w:drawing Element auch ersetzt werden kann:
        // a) [w:docPr, a:graphic, a:graphicData, pic:pic, pic:blip] müssen vorhanden sein
        final Optional<Element> opWpDocPr = elementByTagName("wp:docPr", wDrawingNode);
        final Optional<Element> opAGrapic = elementByTagName("a:graphic", wDrawingNode);
        final Optional<Element> opAGraphicData = elementByTagName("a:graphicData", wDrawingNode);
        final Optional<Element> opPicPic = elementByTagName("pic:pic", wDrawingNode);
        final Optional<Element> opABlip = elementByTagName("a:blip", wDrawingNode);
        
        final boolean neededElementsExist = opWpDocPr.isPresent() && opAGrapic.isPresent()
                && opAGraphicData.isPresent() && opPicPic.isPresent() && opABlip.isPresent();
        
        if (neededElementsExist == false) {
            return;
        }
        
        // b) in w:docPr muss im TITLE der Platzhalter-Name stehen
        final Element wpDocPr = opWpDocPr.get();
        final String keyName = wpDocPr.getAttribute("title");
        
        // Attribut wp:docPr[title] entfernen, da dieser sonst nicht später überschrieben wird
        wpDocPr.removeAttribute("title");
        
        final ImageValue imageValue = getImageValueByKeyName(keyName, values).orElse(null);
        if (null == imageValue) {
            return;
        }
        
        final Object[] registrationResult = registerImageResource(imageValue);
        final String imgRelId = (String) registrationResult[0];
        final boolean usedResourceKeepsExternal = (boolean) registrationResult[1];
        
        setupDrawingElement((Element) wDrawingNode, imageValue, imgRelId, usedResourceKeepsExternal);
    }
    
    private void replacePicture(Node wPictNode, DataValueMap values) {
        // Erstmal schauen ob alles notwendigen Elemente vorhanden sind!
        // [v:shape, v:imagedata]
        final Element vShape = elementByTagName("v:shape", wPictNode).orElse(null);
        final Element vImagedata = elementByTagName("v:imagedata", wPictNode).orElse(null);
        
        if (null == vShape || null == vImagedata) {
            return;
        }
        
        // Dann schauen ob irgendwo ein Platzhalter vergeben wurde
        final String keyName = vShape.getAttribute("alt");
        final ImageValue imageValue = getImageValueByKeyName(keyName, values).orElse(null);
        
        if (null == imageValue) {
            return;
        }
        
        final Object[] registrationResult = registerImageResource(imageValue);
        final String imgRelId = (String) registrationResult[0];
        final boolean usedResourceKeepsExternal = (boolean) registrationResult[1];
        
        setupPictureElement((Element) wPictNode, imageValue, imgRelId);
    }
    
    private Optional<ImageValue> getImageValueByKeyName(String keyName, DataValueMap values) {
        if (null == keyName || keyName.isEmpty() || keyName.trim().isEmpty()) {
            return Optional.empty();
        }
        
        // c) zum Platzhalter muss es einen Wert geben der im Endergebnis einem ImageValue entspr.
        final Optional<DataValue> firstValue = values.getValueByKey(keyName);
        if (firstValue.isPresent() == false) {
            return Optional.empty();
        }
        
        interceptionContext.init(keyName, values);
        final DataValue dataValue = handleInterception(firstValue.get(), interceptionContext);
        if (dataValue.isExtendedValue() == false) {
            return Optional.empty();
        }
        
        final ExtendedValue extendedValue = dataValue.extendedValue();
        if (extendedValue instanceof ImageValue == false) {
            return Optional.empty();
        }
        
        final ImageValue imageValue = (ImageValue) extendedValue;
        
        return Optional.of(imageValue);
    }
    
    /**
     * Registriert anhand des übergebenen {@code imageValue} das Bild sowie die Resource in die
     * notwendigen Dateien und ggf bettet die Bild-Resource direkt ein.
     * 
     * @param imageValue    Im Container zu registrierende und ggf einzubettetes Bild.
     * @return              {@code Object[length 2]}, mit zwei Werten:
     *                      {@code Object[index 0] = String = imageRelationshipId},
     *                      {@code Object[index 1] = boolean = usedResourceKeepsExternal}.
     */
    private Object[] registerImageResource(ImageValue imageValue) {
        final Object[] resultArray = new Object[] {
            "",         // Image-Relationship-Id
            false       // true -> keeps external, false -> is embedded
        };
        
        final ImageResource imageResource = imageValue.getImageResource();
        final ImageResourceType imageType = imageResource.getImageFormatType();
        
        final String imgRelId = "rImgId" + UUID.randomUUID().toString().replace("-", "");
        
        final boolean isExternalResource = imageResource instanceof LocalImageResource
                || imageResource instanceof ExternalImageResource;
        
        final boolean usedResourceKeepsExternal;
        
        if (isExternalResource && loadAndEmbedExternalImages() == false) {
            // Bild-Resource ist extern (Datei || URL) und soll auch nicht geladen werden, sondern
            // im Dokument als externe Resource weiter gereicht werden.
            final String externalTarget;
            if (imageResource instanceof ExternalImageResource) {
                externalTarget = ((ExternalImageResource) imageResource).getResourceURL().toString();
            } else {
                externalTarget = ((LocalImageResource) imageResource).getLocalResource()
                        .toAbsolutePath().toString();
            }
            
            registerRelIdExternalImage(this.extAltChunk.relationshipDocument,
                    externalTarget, imgRelId);
        
            registerContentType(this.extAltChunk.contentTypesDocument,
                    imageType.getMimeType(), imageType.getFileNameExtensions()[0]);
            
            usedResourceKeepsExternal = true;
        } else {
            final String mediaPath = "media/" + imgRelId + "." + imageType.getFileNameExtensions()[0];
            
            extAltChunk.zipFile.createNewFileInZip("word/" + mediaPath);
            extAltChunk.zipFile.overwrite("word/" + mediaPath, imageResource.loadImageData());
            
            registerRelIdEmbeddedImage(this.extAltChunk.relationshipDocument, mediaPath, imgRelId);
            registerContentType(this.extAltChunk.contentTypesDocument,
                    imageType.getMimeType(), imageType.getFileNameExtensions()[0]);
            
            usedResourceKeepsExternal = false;
        }
        
        resultArray[0] = imgRelId;
        resultArray[1] = usedResourceKeepsExternal;
        
        return resultArray;
    }
    
    private void setupDrawingElement(
            Element wDrawingElement, ImageValue imageValue, String imgRelId,
            boolean usedResourceKeepsExternal)
    {
        if (usedResourceKeepsExternal) {
            replaceDrawingRelId(wDrawingElement, imgRelId, true  /* use r:link  */);
        } else {
            replaceDrawingRelId(wDrawingElement, imgRelId, false /* use r:embed */);
        }

        overwriteDrawingElementIds(wDrawingElement);
        applyNonVisiblePropertiesToDrawingElement(wDrawingElement, imageValue);
        applyVisibleTextPropertiesToDrawingElement(wDrawingElement, imageValue);
    }
    
    private void setupPictureElement(
            Element wPictureElement, ImageValue imageValue, String imgRelId)
    {
        replacePictureRelId(wPictureElement, imgRelId);
        overwritePictureElementIds(wPictureElement);

        applyNonVisiblePropertiesToPictureElement(wPictureElement, imageValue);
        applyVisibleTextPropertiesToPictureElement(wPictureElement, imageValue);
    }
    
    private Element createDrawingElement(Document document) {
        final String XML_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        final String XML_NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        final String XML_NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
        
        // w:drawing
        final Element wDrawing = document.createElement("w:drawing");
        
        // w:drawing > wp:inline
        final Element wpInline = (Element) wDrawing.appendChild(document.createElement("wp:inline"));
        wpInline.setAttribute("distR", "0");
        wpInline.setAttribute("distL", "0");
        wpInline.setAttribute("distB", "0");
        wpInline.setAttribute("distT", "0");
        
        final Element wpExtend = (Element) wpInline.appendChild(document.createElement("wp:extent"));
        wpExtend.setAttribute("cy", "357751");
        wpExtend.setAttribute("cx", "573151");
        
        final Element wpEffectExtent = (Element) wpInline.appendChild(document.createElement("wp:effectExtent"));
        wpEffectExtent.setAttribute("r", "2540");
        wpEffectExtent.setAttribute("b", "4445");
        wpEffectExtent.setAttribute("t", "0");
        wpEffectExtent.setAttribute("l", "0");
        
        final Element wpDocPr = (Element) wpInline.appendChild(document.createElement("wp:docPr"));
        wpDocPr.setAttribute("name", "Picture 1"); // wird ggf. überschrieben
        wpDocPr.setAttribute("id", "0"); // ID wird später überschrieben
        
        final Element wpcNvGraphicFramePr = (Element) wpInline.appendChild(document.createElement("wp:cNvGraphicFramePr"));
        final Element aGraphic = (Element) wpInline.appendChild(document.createElement("a:graphic"));
        aGraphic.setAttribute("xmlns:a", XML_NS_A);
        
        // w:drawing > wp:inline > wp:cNvGraphicsFramePt
        final Element aGraphicFrameLocks = (Element) wpcNvGraphicFramePr.appendChild(
                document.createElement("a:graphicFrameLocks"));
        aGraphicFrameLocks.setAttribute("noChangeAspect", "0");
        aGraphicFrameLocks.setAttribute("xmlns:a", XML_NS_A);
        
        // w:drawing > wp:inline > a:graphic
        final Element aGraphicData = (Element) aGraphic.appendChild(document.createElement("a:graphicData"));
        aGraphicData.setAttribute("uri", XML_NS_PIC);
        
        // w:drawing > wp:inline > a:graphic > a:graphicData
        final Element picPic = (Element) aGraphicData.appendChild(document.createElement("pic:pic"));
        picPic.setAttribute("xmlns:pic", XML_NS_PIC);
        
        // w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic
        final Element picNvPicPr = (Element) picPic.appendChild(document.createElement("pic:nvPicPr"));
        final Element picBlipFill = (Element) picPic.appendChild(document.createElement("pic:blipFill"));
        final Element picSpPr = (Element) picPic.appendChild(document.createElement("pic:spPr"));
        picSpPr.setAttribute("bwMode", "auto");
        
        // w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic > pic:nvPicPr
        final Element picCNvPr = (Element) picNvPicPr.appendChild(document.createElement("pic:cNvPr"));
        picCNvPr.setAttribute("name", "Picture 1"); // wird später ggf. überschrieben
        picCNvPr.setAttribute("id", "0"); // ID wird später überschrieben
        
        final Element picCNvPicPr = (Element) picNvPicPr.appendChild(document.createElement("pic:cNvPicPr"));
        final Element apicLocks = (Element) picCNvPicPr.appendChild(document.createElement("a:picLocks"));
        apicLocks.setAttribute("noChangeAspect", "0");
        apicLocks.setAttribute("noChangeArrowheads", "0");
        
        // w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic > pic:blipFill
        final Element aBlip = (Element) picBlipFill.appendChild(document.createElement("a:blip"));
        aBlip.setAttribute("r:embed", "##ERROR##");
        
        final Element aSrcRect = (Element) picBlipFill.appendChild(document.createElement("a:srcRect"));
        final Element aStretch = (Element) picBlipFill.appendChild(document.createElement("a:stretch"));
        
        // w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic > pic:blipFill
        final Element aExtLst = (Element) aBlip.appendChild(document.createElement("a:extLst"));
        
        // w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic > pic:blipFill > a:extLst
        final Element aExt = (Element) aExtLst.appendChild(document.createElement("a:ext"));
        aExt.setAttribute("uri", "{28A0092B-C50C-407E-A947-70E740481C1C}");
        
        // w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic > pic:blipFill > a:extLst > a14:useLocalDpi
        final Element a14UseLocalDpi = (Element) aExt.appendChild(document.createElement("a14:useLocalDpi"));
        a14UseLocalDpi.setAttribute("val", "0");
        a14UseLocalDpi.setAttribute("xmlns:a14", XML_NS_A14);
        
        // w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic > pic:blipFill > a:stretch
        final Element aFillRect = (Element) aStretch.appendChild(document.createElement("a:fillRect"));
        
        // w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic > pic:spPr
        final Element aXfrm = (Element) picSpPr.appendChild(document.createElement("a:xfrm"));
        final Element aPrstGeom = (Element) picSpPr.appendChild(document.createElement("a:prstGeom"));
        aPrstGeom.setAttribute("prst", "rect");
        
        final Element aNoFill = (Element) picSpPr.appendChild(document.createElement("a:noFill"));
        final Element aLn = (Element) picSpPr.appendChild(document.createElement("a:ln"));
        
        // w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic > pic:spPr > a:xfrm
        final Element aOff = (Element) aXfrm.appendChild(document.createElement("a:off"));
        aOff.setAttribute("y", "0");
        aOff.setAttribute("x", "0");
        
        final Element aExt2 = (Element) aXfrm.appendChild(document.createElement("a:ext")); // !! 2 !!
        aExt2.setAttribute("cy", "357751");
        aExt2.setAttribute("cx", "573151");
        
        // w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic > pic:spPr > a:prstGeom
        final Element aAvLst = (Element) aPrstGeom.appendChild(document.createElement("a:avLst"));
        
        // w:drawing > wp:inline > a:graphic > a:graphicData > pic:pic > pic:spPr > a:ln
        final Element aNoFill2 = (Element) aLn.appendChild(document.createElement("a:noFill"));
        
        return wDrawing;
    }
    
    private void replaceDrawingRelId(Element drawingElement, String useRelId, boolean extLink) {
        final String ATTR_R_EMBED = "r:embed";
        final String ATTR_R_LINK = "r:link";
        
        final Element aBlip = elementByTagName("a:blip", drawingElement).orElse(null);
        if (null == aBlip) return;
        
        if (extLink) {
            aBlip.removeAttribute(ATTR_R_EMBED);
            aBlip.setAttribute(ATTR_R_LINK, useRelId);
        } else {
            aBlip.removeAttribute(ATTR_R_LINK);
            aBlip.setAttribute(ATTR_R_EMBED, useRelId);
        }
    }
    
    private void overwriteDrawingElementIds(Element drawingElement) {
        final String myPictureIdMinusOne = Integer.toString(imageCounter);
        final String myPictureId = Integer.toString(++imageCounter);
        
        final Optional<Element> wpDocPr = elementByTagName("wp:docPr", drawingElement);
        if (wpDocPr.isPresent()) {
            wpDocPr.get().setAttribute("name", "Picture " + myPictureId);
            wpDocPr.get().setAttribute("id", myPictureId);
        }
        
        final Optional<Element> picCNvPr = elementByTagName("pic:cNvPr", drawingElement);
        if (picCNvPr.isPresent()) {
            picCNvPr.get().setAttribute("name", "Picture " + myPictureId);
            picCNvPr.get().setAttribute("id", myPictureIdMinusOne);
        }
    }
    
    private void applyVisibleTextPropertiesToDrawingElement(Element wDrawing, ImageValue imgValue) {
        final Optional<Element> opWpDocPr = elementByTagName("wp:docPr", wDrawing);
        if (opWpDocPr.isPresent() == false) {
            return; // Dann lassen wir das vorerst lieber
        }
        
        final Element wpDocPr = opWpDocPr.get();
        
        final String prTitle = imgValue.getTitle().orElse(wpDocPr.getAttribute("title"));
        final String prDescr = imgValue.getDescription().orElse(wpDocPr.getAttribute("desc"));
        
        if (null == prTitle || "".equals(prTitle)) {
            wpDocPr.removeAttribute("title");
        } else {
            wpDocPr.setAttribute("title", prTitle);
        }
        
        if (null == prDescr || "".equals(prDescr)) {
            wpDocPr.removeAttribute("descr");
        } else {
            wpDocPr.setAttribute("descr", prDescr);
        }
    }
    
    private void applyNonVisiblePropertiesToDrawingElement(Element wDrawing, ImageValue imgValue) {
        // Vorher bitte #overwriteDrawingElementIds verwendet und aufrufen
        
        // Überschreibt die nicht-visuellen (also in Word nicht angezeigten) Eigenschaften oder
        // wenn bisher keine dort standen, _versucht_ diese dort einzufügen.
        // Nicht vorhandene Eigenschaften, also jene die in imgValue nie gesetzt wurden, werden
        // nicht überschrieben und auch nicht angefasst, default bleibt erhalten
        
        final Optional<Element> opPicCNvPr = elementByTagName("pic:cNvPr", wDrawing);
        final Element picCNvPr;
        
        if (opPicCNvPr.isPresent()) {
            picCNvPr = opPicCNvPr.get();
        } else {
            // versuchen, anzulegen und einzuhängen, aber nur wenn pic:nvPicPr vorhanden ist
            final Optional<Element> opPicNvPicPr = elementByTagName("pic:nvPicPr", wDrawing);
            if (opPicNvPicPr.isPresent()) {
                picCNvPr = wDrawing.getOwnerDocument().createElement("pic:cNvPr");
                if (opPicCNvPr.get().getChildNodes().getLength() > 0) {
                    opPicCNvPr.get().insertBefore(picCNvPr, opPicCNvPr.get().getFirstChild());
                } else {
                    opPicCNvPr.get().appendChild(picCNvPr);
                }
            } else {
                // das wird nix
                return;
            }
        }
        
        final String orAttrId = picCNvPr.getAttribute("id");
        
        @Nullable String prId = orAttrId.isEmpty() ? Integer.toString(imageCounter - 1) : orAttrId;
        @Nullable String prTitle = imgValue.getTitle().orElse(picCNvPr.getAttribute("title"));
        @Nullable String prDescr = imgValue.getDescription().orElse(picCNvPr.getAttribute("descr"));
        @Nullable String prName = picCNvPr.getAttribute("name");
        
        // Wenn prTitle oder prDescr NULL sind, dann versuche von Lokalen/Externen-Resourcen noch
        // Werte ableiten zu können -> diese werden dann eingesetzt. Ansonsten werden die
        // betreffenden Attribute schlicht entfernt.
        final ImageResource res = imgValue.getImageResource();
        if (    // okay, einer von beiden muss min. NULL sein und die Resource muss extern sein
                (null == prTitle || null == prDescr || "".equals(prTitle) || "".equals(prDescr))
                && (res instanceof LocalImageResource || res instanceof ExternalImageResource))
        {
            if (null == prTitle || "".equals(prTitle)) {
                final String extName = (res instanceof LocalImageResource)
                        ? ((LocalImageResource) res).getLocalResource().getFileName().toString()
                        : ((ExternalImageResource) res).getResourceURL().getHost();
                
                prTitle = extName;
            }
            
            if (null == prDescr || "".equals(prDescr)) {
                final String extPath = (res instanceof LocalImageResource)
                        ? ((LocalImageResource) res).getLocalResource().toString()
                        : ((ExternalImageResource) res).getResourceURL().toString();
                
                prDescr = extPath;
            }
            
            if ((null != prName && prName.startsWith("Picture")) && res instanceof LocalImageResource) {
                prName = ((LocalImageResource) res).getLocalResource().getFileName().toString();
            }
        }
        
        picCNvPr.setAttribute("id", prId);
        
        if (null == prTitle || "".equals(prTitle)) {
            picCNvPr.removeAttribute("title");
        } else {
            picCNvPr.setAttribute("title", prTitle);
        }
        
        if (null == prDescr || "".equals(prDescr)) {
            picCNvPr.removeAttribute("descr");
        } else {
            picCNvPr.setAttribute("descr", prDescr);
        }
        
        picCNvPr.setAttribute("name", (null == prName) ? "" : prName);
    }
    
    private Element createPictureElement(Document document) {
        final Element wPict = document.createElement("w:pict");
        final Element vShape = (Element) wPict.appendChild(document.createElement("v:shape"));
        vShape.setAttribute("id", "vShapeImage0");
        vShape.setAttribute("type", "#_x0000_t75");
        vShape.setAttribute("style", "width:auto; height:auto");
        
        final Element vImagedata = (Element) vShape.appendChild(document.createElement("v:imagedata"));
        vImagedata.setAttribute("r:id", "##ERROR##");
        
        return wPict;
    }
    
    private void replacePictureRelId(Element pictureElement, String useRelId) {
        final Element vImageData = (Element) pictureElement.getElementsByTagName("v:imagedata").item(0);
        vImageData.setAttribute("r:id", useRelId);
    }
    
    private void overwritePictureElementIds(Element pictureElement) {
        final Optional<Element> vShape = elementByTagName("v:shape", pictureElement);
        if (vShape.isPresent()) {
            vShape.get().setAttribute("id", "vShapeImage" + Integer.toString(++imageCounter));
        }
    }
    
    private void applyVisibleTextPropertiesToPictureElement(Element wPict, ImageValue imgValue) {
        // Überschreit lediglich ein Attribut v:shape[alt]
        final Element vShape = elementByTagName("v:shape", wPict).orElse(null);
        if (null == vShape) return;
        
        final String prTitle = imgValue.getTitle().orElse("");
        final String prDescr = imgValue.getDescription().orElse("");
        
        if (prTitle.isEmpty() && prDescr.isEmpty()) {
            // Attribut 'alt' fliegt dann einfach raus
            vShape.removeAttribute("alt");
        } else if (prTitle.isEmpty() && prDescr.isEmpty() == false) {
            vShape.setAttribute("alt", prDescr);
        } else if (prTitle.isEmpty() == false && prDescr.isEmpty()) {
            vShape.setAttribute("alt", prTitle);
        } else if (prTitle.isEmpty() == false && prDescr.isEmpty() == false) {
            vShape.setAttribute("alt", prTitle + "\n" + prDescr);
        }
    }
    
    private void applyNonVisiblePropertiesToPictureElement(Element wPict, ImageValue imgValue) {
        final Element vImagedata = elementByTagName("v:imagedata", wPict).orElse(null);
        if (null == vImagedata) return;
        
        // v:imagedata[o:title]
        final String prTitle = imgValue.getTitle().orElse(null);
        if (null == prTitle || prTitle.isEmpty()) {
            vImagedata.removeAttribute("o:title");
        } else {
            vImagedata.setAttribute("o:title", prTitle);
        }
        
//        // v:imagedata[o:althref]
//        final ImageResource imgRes = imgValue.getImageResource();
//        if (imgRes instanceof LocalImageResource || imgRes instanceof ExternalImageResource) {
//            final String prOHref;
//            
//            if (imgRes instanceof LocalImageResource) {
//                prOHref = ((LocalImageResource) imgRes).getLocalResource().toString();
//            } else {
//                prOHref = ((ExternalImageResource) imgRes).getResourceURL().toString();
//            }
//            
//            vImagedata.setAttribute("o:althref", prOHref);
//        } else {
//            vImagedata.removeAttribute("o:althref");
//        }
    }
    
    private Element createRelationshipImageEmbeddedElement(Document document, String imageTarget, String rId) {
        final String IMG_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        
        final Element relationship = (Element) document.createElement("Relationship");
        relationship.setAttribute("Id", rId);
        relationship.setAttribute("Type", IMG_TYPE);
        relationship.setAttribute("Target", imageTarget);
        
        return relationship;
    }
    
    private Element createRelationshipImageExternalElement(Document document, String imageTarget, String rId) {
        final Element relElement = createRelationshipImageEmbeddedElement(document, imageTarget, rId);
        relElement.setAttribute("TargetMode", "External");
        return relElement;
    }
    
    private void registerRelIdEmbeddedImage(Document document, String imageTarget, String rId) {
        registerRelIdForImage0(document, imageTarget, rId, true /* ja, erstelle eingebettet */);
    }
    
    private void registerRelIdExternalImage(Document document, String imageTarget, String rId) {
        registerRelIdForImage0(document, imageTarget, rId, false /* nein, nicht eingebettet, extern */);
    }
    
    private void registerRelIdForImage0(Document document, String imageTarget, String rId, boolean createEmbeddedId) {
        // TODO: nicht einfach von Index 0 nehmen, sondern prüfen ob dort überhaupt ein Element ist!
        final Element rootRelElement = (Element) document.getElementsByTagName("Relationships").item(0);
        final NodeList relationshipList = rootRelElement.getChildNodes();
        
        for (int nodeIndex = 0; nodeIndex < relationshipList.getLength(); nodeIndex++) {
            final Node nodeItem = relationshipList.item(nodeIndex);
            if (nodeItem instanceof Element == false) continue;
            
            final Element relElement = (Element) nodeItem;
            if (relElement.getNodeName().equals("Relationship") == false) continue;
            
            final String attrRId = relElement.getAttribute("Id");
            
            if (null != attrRId && attrRId.equals(rId)) {
                // dann wurde diese ID bereits vormals erfolgreich registriert und muss nicht
                // (darf nicht) doppelt eingetragen werden.
                return;
            }
        }
        
        // Okay, ID unbekannt -> registrieren
        final Element relImgElement = createEmbeddedId
                ? createRelationshipImageEmbeddedElement(document, imageTarget, rId)
                : createRelationshipImageExternalElement(document, imageTarget, rId);
        
        if (relationshipList.getLength() == 0) {
            rootRelElement.appendChild(relImgElement);
        } else {
            rootRelElement.insertBefore(relImgElement, rootRelElement.getFirstChild());
        }
    }
    
    private Element createContentTypeElement(Document document, String mimeContentType, String extension) {
        final Element xmlDefault = (Element) document.createElement("Default");
        xmlDefault.setAttribute("ContentType", mimeContentType);
        xmlDefault.setAttribute("Extension", extension);
        return xmlDefault;
    }
    
    private void registerContentType(Document document, String mimeContentType, String extension) {
        final Element rootTypesElement = (Element) document.getElementsByTagName("Types").item(0);
        final NodeList contentTypeList = rootTypesElement.getChildNodes();
        
        for (int nodeIndex = 0; nodeIndex < contentTypeList.getLength(); nodeIndex++) {
            final Node nodeItem = contentTypeList.item(nodeIndex);
            if (nodeItem instanceof Element == false) continue;
            
            final Element typeElement = (Element) nodeItem;
            if (typeElement.getNodeName().equals("Default") == false) continue;
            
            final String attrContentType = typeElement.getAttribute("ContentType");
            if (null != attrContentType && attrContentType.equals(mimeContentType)) {
                // MIME Type ist bereits enthalten und registriert
                return;
            }
        }
        
        // Noch nicht registriert, dann einfach hinzufügen
        final Element newMimeType = createContentTypeElement(document, mimeContentType, extension);
        
        if (contentTypeList.getLength() == 0) {
            rootTypesElement.appendChild(newMimeType);
        } else {
            rootTypesElement.insertBefore(newMimeType, rootTypesElement.getFirstChild());
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
    
    private GenericNodeIterator walkDrawingButNoTables(Node rootNode) {
        return new GenericNodeIterator(rootNode, "w:drawing")
                .noRecursionByElements("w:tbl")
                .doSkipFirstStopElement();
    }
    
    private GenericNodeIterator walkPictureButNoTables(Node rootNode) {
        return new GenericNodeIterator(rootNode, "w:pict")
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
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // IMPLEMENTIERUNG DER MS ERWEITERUNG FÜR ALTERNATIVE CHUNK ELEMENTE/PLATZHALTER
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    /**
     * Unterklasse der Erweiterung um w:altChunk mit der Implementierung für ZIP Container.
     */
    private class InnerAltChunkExtension extends MicrosoftInsertDoc {
        
        private ZIPDocumentFile zipFile;
        private Document wordDocument;
        private Document relationshipDocument;
        private Document contentTypesDocument;
        
        public InnerAltChunkExtension setCurrentZipFile(ZIPDocumentFile zipFile) {
            this.zipFile = zipFile;
            return this;
        }
        
        public InnerAltChunkExtension setWordDocument(Document doc) {
            this.wordDocument = doc;
            return this;
        }
        
        public InnerAltChunkExtension setRelationshipDocument(Document doc) {
            this.relationshipDocument = doc;
            return this;
        }

        @Override
        protected Document getWordDocument() {
            return wordDocument;
        }

        @Override
        protected Document getRelationshipDocument() {
            return relationshipDocument;
        }

        @Override
        protected Document getContentTypeDocument() {
            return this.contentTypesDocument;
        }

        @Override
        protected void overwritePartInContainer(String partName, byte[] data) {
            assert null != partName : "partName == null";
            assert null != data : "data == null";
            assert null != zipFile : "#setCurrentZipFile(null??)";
            
            zipFile.createNewFileInZip(partName);
            zipFile.overwrite(partName, data);
        }

        private InnerAltChunkExtension setContentTypesDocument(Document newContentTypes) {
            this.contentTypesDocument = newContentTypes;
            return this;
        }
        
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // INTERCEPTION-CONTEXT FÜR MS DOKUMENTE
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private class MyInterceptionContext extends InterceptionContext {
        
        private String placeholder;
        private DataValueMap<?> valueMap;
        
        public void init(String placeholder, DataValueMap<?> valueMap) {
            this.placeholder = placeholder;
            this.valueMap = valueMap;
        }
        
        ////////////////////////////////////////////////////////////////////////////////////////////

        @Override
        public OfficeDocumentFactory getDocumentFactory() {
            return MicrosoftDocument.this.getRelatedFactory();
        }

        @Override
        public OfficeDocument getDocument() {
            return MicrosoftDocument.this;
        }

        @Override
        public String getPlaceholderName() {
            return placeholder;
        }

        @Override
        public DataValueMap<?> getParentValueMap() {
            return valueMap;
        }

        @Override
        public boolean isXmlBasedDocument() {
            return true;
        }
        
    }

}
