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
import com.mz.solutions.office.model.DataMap;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataTable;
import com.mz.solutions.office.model.DataTableRow;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.model.DataValueMap;
import com.mz.solutions.office.model.ValueOptions;
import com.mz.solutions.office.model.interceptor.InterceptionContext;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import static mz.solutions.office.resources.MessageResources.formatMessage;
import static mz.solutions.office.resources.OpenDocumentKeys.NO_DATA;
import static mz.solutions.office.resources.OpenDocumentKeys.UNKNOWN_FORMATTING_CHAR;
import static mz.solutions.office.resources.OpenDocumentKeys.UNKNOWN_PLACE_HOLDER;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

final class OpenDocument extends AbstractOfficeXmlDocument {
    
    private static final String ZIP_DOC_CONTENT = "content.xml";
    private static final String ZIP_DOC_STYLES = "styles.xml";
    
    private final MyInterceptionContext interceptionContext = new MyInterceptionContext();
    
    public OpenDocument(OfficeDocumentFactory factory, Path document) {
        super(factory, document, ZIP_DOC_CONTENT, ZIP_DOC_STYLES);
    }

    @Override
    protected String getImplementedOfficeName() {
        return "Apache OpenOffice 4.x / LibreOffice";
    }

    
    @Override
    protected ZIPDocumentFile createAndFillDocument(
            final Iterator<DataPage> dataPages) {
        
        final ZIPDocumentFile newFile = sourceDocumentFile.cloneDocument();
        
        fillDocuments0(dataPages, newFile);
            
        return newFile;
    }
    
    private void fillDocuments0(
            final Iterator<DataPage> dataPages,
            final ZIPDocumentFile outputDocument) {
        
        final Document newContent = (Document) sourceContent.cloneNode(true);
        final Document newStyles = (Document) sourceStyles.cloneNode(true);
        
        removeUserFieldDeclaration(newContent);
        
        final Node nodeContentBody = findDocumentBody(newContent);
        final Node nodeContentBodyCopy = nodeContentBody.cloneNode(true);
        
        // Leeren Content-Body um alle Datensätze anfügen zu können
        final Node newFullContentBody = nodeContentBody.cloneNode(true);
        final NodeList bodyChildNodes = newFullContentBody.getChildNodes();
        
        for (int i = 0; i < bodyChildNodes.getLength(); i++) {
            newFullContentBody.removeChild(bodyChildNodes.item(i));
        }
        
        boolean firstPage = true;
        boolean missingDataPages = true;
        
        while (dataPages.hasNext()) {
            final DataPage nextPage = dataPages.next();
            final Node newContentBody = nodeContentBody.cloneNode(true);
            
            replaceFields(newContentBody, nextPage);
            
            // Alle Elemente im leeren Content-Body anfügen
            final NodeList filledNodes = newContentBody.getChildNodes();
            
            for (int i = 0; i < filledNodes.getLength(); i++) {
                newFullContentBody.appendChild(filledNodes.item(i));
            }
            
            missingDataPages = false;
            firstPage = false;
        }
        
        if (missingDataPages) {
            if (ignoreMissingDataPages()) {
                return; // Leise beenden
            }
            
            throw new NoDataForDocumentGenerationException(
                    formatMessage(NO_DATA));
        }
        
        // Gefüllten Body-Node tauschen mit ursprünglichem Bode-Node
        final Node parentNode = nodeContentBody.getParentNode();
        parentNode.replaceChild(newFullContentBody, nodeContentBody);

        overwrite(outputDocument, ZIP_DOC_CONTENT, newContent);
        overwrite(outputDocument, ZIP_DOC_STYLES, newStyles);
    }
    
    private void removeUserFieldDeclaration(Node documentRoot) {
        final Node textUserFieldDecls =  findUserFieldDeclaration(documentRoot);
        
        if (null == textUserFieldDecls) {
            return; // Keine Benutzerdefinierten Felder
        }
        
        final Node parentNode = textUserFieldDecls.getParentNode();
        
        parentNode.removeChild(textUserFieldDecls);
    }
    
    private void replaceFields(Node node, DataMap valueMap) {
        // Normale Platzhalter ersetzen (nicht rekursiv Tabellen)
        replaceUserFields(node, valueMap);
        
        // Inhalt von Tabellen die unbekannt sind, dann eben auch normal
        // mit den vorhandenen Platzhaltern ersetzen
        // >> bekannte Tabellen werden nicht ersetzt
        replaceUnkownTables(node, valueMap);
        
        // Tabellen die bekannt sind, müssen speziell hier behandelt werden
        for (Node tableNode : walkTables(node)) {
            final String tableName = getTableName(tableNode);
            final boolean isKnownTable = valueMap
                    .getTableByName(tableName)
                    .isPresent();

            if (isKnownTable) {
                replaceKnownTable(tableNode, valueMap);
            }
        }
    }
    
    private void replaceUserFields(Node documentBody, DataValueMap value) {
        for (Node userFieldNode : walkUserFields(documentBody)) {
            replaceUserFieldNode(userFieldNode, value);
        }
    }
    
    private void replaceUnkownTables(Node node, DataMap valueMap) {
        for (Node tableNode : walkTables(node)) {
            final String name = getTableName(tableNode);
            final Optional<DataTable> table = valueMap.getTableByName(name);
            
            // Die Tabelle darf NUR ersetzt werden, wenn es eine unbekannte
            // Tabelle ist; für die Ersetzung von Tabellen mit Datenreihen
            // ist diese Methode nicht zuständig und muss entsprechende
            // Tabellen überspringen!
            
            if (table.isPresent() == true) {
                continue; // ignorieren
            }
            
            replaceUserFields(tableNode, valueMap);
        }
    }
    
    private void replaceKnownTable(Node tableNode, DataMap values) {
        final String tableName = getTableName(tableNode);
        final Optional<DataTable> tableData = values.getTableByName(tableName);
        
        if (tableData.isPresent() == false) {
            // DARF EIGENTLICH NICHT MEHR AUFTRETEN!
            // Aufgrund der Ausnahmesituation wird hier auf
            // eine Übersetzung verzichtet.
            throw new IllegalStateException(
                    "(Internal Error) Found unknown table: " + tableName);
        }
        
        final List<Node> tableRows = walkTableRows(tableNode).asList();
        
        final int dataRowIndex = tableRows.size() == 1 ? 0 : 1;
        final Node tableDataRow = tableRows.get(dataRowIndex);
        
        // Alle Zeilen die NICHT die zu wiederholende Datenzeilen sind
        // ganz normal ersetzen
        for (int i = 0; i < tableRows.size(); i++) {
            if (dataRowIndex == i /* current */) {
                continue;
            }
            
            replaceUserFields(tableRows.get(i), tableData.get());
        }
        
        final Iterator<DataTableRow> rowIterator = tableData.get().iterator();
        while (rowIterator.hasNext()) {
            final DataTableRow rowData = rowIterator.next();
            final Node newTableRow = tableDataRow.cloneNode(true);
            
            replaceFields(newTableRow, rowData);
            
            tableNode.insertBefore(newTableRow, tableDataRow);
        }
        
        tableNode.removeChild(tableDataRow);
    }
    
    private void replaceUserFieldNode(Node userFieldNode, DataValueMap values) {
        final String fieldName = getFieldName(userFieldNode).trim();
        final Optional<DataValue> value = values.getValueByKey(fieldName);
        
        if (value.isPresent() == false) {
            if (ignoreMissingValues() == true) {
                return;
            }
            
            throw new DocumentPlaceholderMissingException(
                    formatMessage(UNKNOWN_PLACE_HOLDER,
                            /* {0} */ fieldName));
        }
        
        interceptionContext.init(fieldName, values);
        
        final DataValue dataValue = handleInterception(value.get(), interceptionContext);
        final Set<ValueOptions> options = dataValue.getValueOptions();
        
        final boolean isSimpleText =
                options.contains(ValueOptions.KEEP_LINEBREAK) == false
                && options.contains(ValueOptions.KEEP_TABULATOR) == false;
        
        final String textContent = dataValue.getValue();
        
        final Document document = userFieldNode.getOwnerDocument();
        final Node parentNode = userFieldNode.getParentNode();
        
        if (isSimpleText) {
            final Node textNode = document.createTextNode(textContent);
            
            parentNode.replaceChild(textNode, userFieldNode);
            return;
        }
        
        final StringBuilder textBuilder = new StringBuilder();
        
        for (char chLetter : textContent.toCharArray()) {
            if (chLetter == '\r') {
                continue;
            }
            
            final boolean special = chLetter == '\n' || chLetter == '\t';
            
            if (special == false) {
                textBuilder.append(chLetter);
                continue;
            }
            
            // Behandlung Formatierungsangabe
            final boolean needTextNode = textBuilder.length() != 0;
            
            if (needTextNode) {
                final String nodeText = textBuilder.toString();
                final Node textNode = document.createTextNode(nodeText);
                
                parentNode.insertBefore(textNode, userFieldNode);
                textBuilder.setLength(0);
            }
            
            final Node formatElement;
            
            switch (chLetter) {
                case '\n': formatElement = document.createElement("text:line-break"); break;
                case '\t': formatElement = document.createElement("text:tab"); break;
                default:
                    throw new IllegalStateException(
                            formatMessage(UNKNOWN_FORMATTING_CHAR,
                                    /* {0} */ Integer.toString((int) chLetter)));
            }
            
            parentNode.insertBefore(formatElement, userFieldNode);
        }
        
        if (textBuilder.length() > 0) {
            final String text = textBuilder.toString();
            final Node textNode = document.createTextNode(text);
            
            parentNode.insertBefore(textNode, userFieldNode);
        }
        
        parentNode.removeChild(userFieldNode);
    }
    
    ////////////////////////////////////////////////////////////////////////////
    // XML Manipulations Methoden
    ////////////////////////////////////////////////////////////////////////////

    private Node findDocumentBody(Node rootNode) {
        return findNodeByName(rootNode, "office:body");
    }
    
    private Node findUserFieldDeclaration(Node rootNode) {
        return findNodeByName(rootNode, "text:user-field-decls");
    }
    
    private String getFieldName(Node fieldNode) {
        return getAttribute(fieldNode, "text:name");
    }
    
    private String getTableName(Node tableNode) {
        return getAttribute(tableNode, "table:name");
    }
    
    private GenericNodeIterator walkTables(Node rootNode) {
        return new GenericNodeIterator(rootNode, "table:table");
    }
    
    private GenericNodeIterator walkTableRows(Node tableNode) {
        return new GenericNodeIterator(tableNode, "table:table-row")
                .noRecursionByElements("table:table-cell");
    }
    
    private GenericNodeIterator walkUserFields(Node rootNode) {
        return new GenericNodeIterator(rootNode, "text:user-field-get")
                .noRecursionByElements("table:table");
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // INTERCEPTION CONTEXT FÜR DIE VERWENDUNG VON VALUE-INTERCEPTOR'S
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private class MyInterceptionContext extends InterceptionContext {

        private String placeholder;
        private DataValueMap<?> valueMap;
        
        public void init(String placeholder, DataValueMap<?> valueMap) {
            this.placeholder = placeholder;
            this.valueMap = valueMap;
        }
        
        @Override
        public OfficeDocumentFactory getDocumentFactory() {
            return OpenDocument.this.getRelatedFactory();
        }

        @Override
        public OfficeDocument getDocument() {
            return OpenDocument.this;
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
