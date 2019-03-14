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
package com.mz.solutions.office;

import com.mz.solutions.office.OfficeDocumentException.DocumentPlaceholderMissingException;
import com.mz.solutions.office.OfficeDocumentException.NoDataForDocumentGenerationException;
import com.mz.solutions.office.extension.ExtendedValue;
import com.mz.solutions.office.instruction.DocumentInterceptor;
import com.mz.solutions.office.model.DataMap;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataTable;
import com.mz.solutions.office.model.DataTableRow;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.model.DataValueMap;
import com.mz.solutions.office.model.ValueOptions;
import com.mz.solutions.office.model.hints.StandardFormatHint;
import com.mz.solutions.office.model.images.ExternalImageResource;
import com.mz.solutions.office.model.images.ImageResource;
import com.mz.solutions.office.model.images.ImageResourceType;
import com.mz.solutions.office.model.images.ImageValue;
import com.mz.solutions.office.model.images.LocalImageResource;
import com.mz.solutions.office.model.interceptor.InterceptionContext;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Collections;
import java.util.IdentityHashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.UUID;
import static mz.solutions.office.resources.MessageResources.formatMessage;
import static mz.solutions.office.resources.OpenDocumentKeys.NO_DATA;
import static mz.solutions.office.resources.OpenDocumentKeys.UNKNOWN_FORMATTING_CHAR;
import static mz.solutions.office.resources.OpenDocumentKeys.UNKNOWN_PLACE_HOLDER;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

final class OpenDocument extends AbstractOfficeXmlDocument {
    
    private static final String ZIP_DOC_CONTENT = "content.xml";
    private static final String ZIP_DOC_STYLES = "styles.xml";
    private static final String ZIP_MANIFEST = "META-INF/manifest.xml";
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    /**
     * Eltern-Elemente welche als Absatz nach OASIS ODF 1-2 zu interpretieren sind.
     * <p>Sowohl {@code text:p} wie auch {@code text:h} stellen Absatz-Elemente dar.</p>
     */
    private static final List<String> OASIS_ODF_PARAGRAPH_ELEMENTS = Collections.unmodifiableList(
            Arrays.asList("text:p", "text:h"));
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private final MyInterceptionContext interceptionContext = new MyInterceptionContext();
    
    private volatile int imageCounter = 16_000;
    private volatile Map<ImageResource, String> cacheImageResources = new IdentityHashMap<>();
    
    public OpenDocument(OfficeDocumentFactory factory, Path document) {
        super(factory, document);
    }

    @Override
    protected String getImplementedOfficeName() {
        return "Apache OpenOffice 4.x / LibreOffice";
    }
    
    @Override
    protected void createAndFillDocument(Iterator<DataPage> dataPages) {
        this.imageCounter = 16_000;
        this.cacheImageResources.clear();
        
        try {
            fillDocuments0(dataPages);
        } finally {
            this.cacheImageResources.clear();
        }
    }
    
    private void fillDocuments0(final Iterator<DataPage> dataPageIterator) {
        final Document documentContent = getDocumentPart(ZIP_DOC_CONTENT);
        removeUserFieldDeclaration(documentContent);
        
        final Node nodeContentBody = findDocumentBody(documentContent);
        
        // Leeren Content-Body um alle Datensätze anfügen zu können
        final Node newFullContentBody = nodeContentBody.cloneNode(true);
        final NodeList bodyChildNodes = newFullContentBody.getChildNodes();
        
        for (int i = 0; i < bodyChildNodes.getLength(); i++) {
            newFullContentBody.removeChild(bodyChildNodes.item(i));
        }
        
        boolean missingDataPages = true;
        
        while (dataPageIterator.hasNext()) {
            final DataPage nextPage = dataPageIterator.next();
            final Node newContentBody = nodeContentBody.cloneNode(true);
            
            replaceDocumentTree(newContentBody, nextPage);
            
            // Alle Elemente im leeren Content-Body anfügen
            final NodeList filledNodes = newContentBody.getChildNodes();
            
            for (int i = 0; i < filledNodes.getLength(); i++) {
                newFullContentBody.appendChild(filledNodes.item(i));
            }
            
            missingDataPages = false;
        }
        
        processHeaderFooter();
        
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
    }
    
    private Document getDocumentManifest() {
        return getDocumentPart(ZIP_MANIFEST);
    }
    
    private void processHeaderFooter() {
        if (hasHeaderFooterInstructions() == false) return;
        
        final Document stylesDocument = getDocumentPart(ZIP_DOC_STYLES);
        final Optional<Element> officeMasterStyles = elementByTagName(
                "office:master-styles", stylesDocument.getDocumentElement());
        
        if (officeMasterStyles.isPresent() == false) return;
        
        final NodeList headerElements = officeMasterStyles.get().getElementsByTagName("style:header");
        final NodeList footerElements = officeMasterStyles.get().getElementsByTagName("style:footer");
        
        final String attrPageStyleName = officeMasterStyles.get().getAttribute("style:name");
        final String attrPageLayoutName = officeMasterStyles.get().getAttribute("style:page-layout-name");
        
        for (int i = 0; i < headerElements.getLength(); i++) {
            replaceHeaderFooterElement(
                    (Element) headerElements.item(i), true /* header */,
                    attrPageStyleName, attrPageLayoutName);
        }
        
        for (int i = 0; i < footerElements.getLength(); i++) {
            replaceHeaderFooterElement(
                    (Element) footerElements.item(i), false /* header */,
                    attrPageStyleName, attrPageLayoutName);
        }
    }
    
    private void replaceHeaderFooterElement(Element element, boolean header, String ... names) {
        DataMap<?> values = null;
        
        for (String singleName : names) {
            if (null == singleName) continue;
            if (singleName.isEmpty()) continue;
            
            final Optional<DataMap<?>> opData = callHeaderFooterInstruction(singleName, header);
            
            if (null != opData && opData.isPresent()) {
                values = opData.get();
                break;
            }
        }
        
        if (null == values) {
            Optional<DataMap<?>> opData = callHeaderFooterInstruction("", header);
            if (null != opData && opData.isPresent()) {
                values = opData.get();
            }
        }
        
        if (null == values) return; // Kein Ersetzungsvorgang
        
        replaceDocumentTree(element, values);
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    @Override
    protected String normalizePartName(String inputPartName) {
        if (DocumentInterceptor.GENERIC_PART_BODY.equals(inputPartName)) {
            return ZIP_DOC_CONTENT;
        }
        
        if (DocumentInterceptor.GENERIC_PART_STYLES.equals(inputPartName)) {
            return ZIP_DOC_STYLES;
        }
        
        return inputPartName;
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private void removeUserFieldDeclaration(Node documentRoot) {
        final Node textUserFieldDecls =  findUserFieldDeclaration(documentRoot);
        
        if (null == textUserFieldDecls) {
            return; // Keine Benutzerdefinierten Felder
        }
        
        final Node parentNode = textUserFieldDecls.getParentNode();
        
        parentNode.removeChild(textUserFieldDecls);
    }
    
    private void replaceDocumentTree(Node node, DataMap valueMap) {
        replaceFieldsAndImages(node, valueMap);
        replaceTables(node, valueMap);
    }
    
    private void replaceTables(Node node, DataMap valueMap) {
        for (Node tableNode : walkTables(node)) {
            final String tableName = getTableName(tableNode);
            final Optional<DataTable> table = valueMap.getTableByName(tableName);
            
            final boolean tableIsKnown = table.isPresent();
            
            if (tableIsKnown) {
                replaceKnownTable(tableNode, valueMap);
            } else {
                replaceUnknownTable(tableNode, valueMap);
            }
        }
    }
    
    private void replaceFieldsAndImages(Node documentBody, DataValueMap value) {
        for (Node drawFrameNode : walkDrawFrames(documentBody)) {
            replaceDrawFrame((Element) drawFrameNode, value);
        }
        
        for (Node userFieldNode : walkUserFields(documentBody)) {
            replaceFieldNode(userFieldNode, value);
        }
        
        for (Node placeholderNode : walkPlaceholders(documentBody)) {
            replaceFieldNode(placeholderNode, value);
        }
    }
    
    private void replaceUnknownTable(Node tableNode, DataMap values) {
        final NodeList childNodes = tableNode.getChildNodes();
        for (Node anyNode : toFlatNodeList(childNodes)) {
            replaceDocumentTree(anyNode, values);
        }
    }
    
    private void replaceKnownTable(Node tableNode, DataMap values) {
        final String tableName = getTableName(tableNode);
        final Optional<DataTable> tableData = values.getTableByName(tableName);
        
        if (tableData.isPresent() == false) {
            throw new IllegalStateException(
                    "(Internal Error) Found unknown table: " + tableName);
        }
        
        final List<Node> tableRows = walkTableRows(tableNode).asList();
        
        final int dataRowIndex;// = tableRows.size() - 1; // tableRows.size() == 1 ? 0 : 1;
        final boolean thereIsNoDataRow;
        
        if (hasTableHeaderRows(tableNode)) {
            int assumedDataRowIndex = -1; // Keine Data-Row
            
            for (int tableRowIndex = 0; tableRowIndex < tableRows.size(); tableRowIndex++) {
                final Element tableRow = (Element) tableRows.get(tableRowIndex);
                final Node tableRowParent = tableRow.getParentNode();
                
                if (tableRowParent instanceof Element == false) continue;
                
                final boolean isHeaderRow = ((Element) tableRowParent).getNodeName()
                        .equals("table:table-header-rows");
                
                if (isHeaderRow) {
                    continue;   // Kopf-Zeilen ignorieren wir einfach
                } else {
                    assumedDataRowIndex = tableRowIndex;
                    break;
                }
            }
            
            // Wenn dataRowIndex||assumedDataRowIndex == -1 dann:
            // Keine Datenzeile gefunden? Dann sind alles Kopf-Zeilen. Ersetzungsvorgang
            // kann somit nach dem Ersetzen der Kopfzeilen enden.
            if (assumedDataRowIndex == -1) {
                thereIsNoDataRow = true;
            } else {
                thereIsNoDataRow = false;
            }
            
            dataRowIndex = assumedDataRowIndex;
        } else {
            dataRowIndex = tableRows.size() == 1 ? 0 : 1;
            thereIsNoDataRow = false;
        }
        
        final Node tableDataRow = tableRows.get(dataRowIndex);
        
        // Alle Zeilen die NICHT die zu wiederholende Datenzeilen sind
        // ganz normal ersetzen
        for (int i = 0; i < tableRows.size(); i++) {
            if (dataRowIndex == i /* current */) {
                continue;
            }
            
            replaceFieldsAndImages(tableRows.get(i), tableData.get());
        }
        
        if (thereIsNoDataRow) {
            // Keine Datenzeile vorhanden. Also kann die Ersetzung jener ausfallen.
            removeTableElementIfZeroRows(tableNode);
            return;
        }
        
        final Iterator<DataTableRow> rowIterator = tableData.get().iterator();
        while (rowIterator.hasNext()) {
            final DataTableRow rowData = rowIterator.next();
            final Node newTableRow = tableDataRow.cloneNode(true);
            
            replaceDocumentTree(newTableRow, rowData);
            
            tableNode.insertBefore(newTableRow, tableDataRow);
        }
        
        tableNode.removeChild(tableDataRow);
        removeTableElementIfZeroRows(tableNode);
    }
    
    private boolean hasTableHeaderRows(Node tableNode) {
        final NodeList tableChilds = tableNode.getChildNodes();
        
        if (tableChilds.getLength() == 0)  return false;
        
        for (int i = 0; i < tableChilds.getLength(); i++) {
            final Node tableChildNode = tableChilds.item(i);
            
            if (tableChildNode instanceof Element == false) continue;
            
            if (((Element) tableChildNode).getNodeName().equals("table:table-header-rows")) {
                return true;
            }
        }
        
        return false;
    }
    
    /**
     * Wenn nach dem Ersetzungsvorgang eine Tabelle gar keine Datenzeilen besitzt (also überhaupt
     * keine Zeilen), sowohl keine Kopf wie Fußzeilen, dann erkennt diese Methode dies und entfernt
     * die in Office eh nicht mehr sichtbare Tabelle auf XML Ebene komplett.
     * 
     * @param tableNode     Element {@code table:table}.
     */
    private void removeTableElementIfZeroRows(Node tableNode) {
        final Element tableTable = (Element) tableNode;
        final NodeList tableTableRowElements = tableTable.getElementsByTagName("table:table-row");
        
        if (tableTableRowElements.getLength() == 0) {
            // Dann haben wir ein quasi leeres table:table Element (außergenommen paar anderer
            // Defintionen die nicht sichtbar sind) und jenes Element entfernen wir gleich ganz.
            tableTable.getParentNode().removeChild(tableTable);
        }
    }
    
    private void replaceFieldNode(Node userFieldNode, DataValueMap values) {
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
        if (dataValue.isExtendedValue() && dataValue.extendedValue() instanceof ImageValue) {
            // Ohhh Platzhalter ist ein Bild!
            replaceUserFieldWithDrawFrame(userFieldNode, (ImageValue) dataValue.extendedValue());
            return;
        }
        
        if (dataValue.isExtendedValue() && dataValue.extendedValue() instanceof StandardFormatHint) {
            final StandardFormatHint formatHint = (StandardFormatHint) dataValue.extendedValue();
            
            if (formatHint.equals(StandardFormatHint.PARAGRAPH_KEEP)) {
                // Okay, KEEP ist zwar dar, aber wir müssen nichts machen. Der Absatz soll ja
                // bewusst stehen gelassen werden.
                
                userFieldNode.getParentNode().removeChild(userFieldNode);
                return;
            }
            
            if (formatHint.equals(StandardFormatHint.PARAGRAPH_REMOVE)) {
                // Okay, wir müssen nicht unseren eigenen Platzhalter entfernen. Warum? Weil wir
                // gleich den ganzen Absatz entfernen in dem jener steht. :-D
                
                Node currentNode = userFieldNode;
                do {
                    if (currentNode instanceof Element) {
                        final Element anyElement = (Element) currentNode;
                        final String elementName = anyElement.getNodeName();
                        
                        final boolean isParentParagraphElement =
                                OASIS_ODF_PARAGRAPH_ELEMENTS.contains(elementName);
                        
                        if (isParentParagraphElement) {
                            // Okay, gesuchtes Element IST das oberste Absatz-Element. Das fliegt
                            // raus.
                            anyElement.getParentNode().removeChild(anyElement);
                            return;
                        }
                    }
                    
                    currentNode = currentNode.getParentNode();
                } while (null != currentNode && isBodyElement(currentNode) == false);
                return;
            }
            
            if (formatHint.equals(StandardFormatHint.PARAGRAPH_HIDDEN)) {
                // TODO Absatz als Hidden deklarieren
                markStyleTextOrParagraphHiddenFor((Element) userFieldNode.getParentNode());
                userFieldNode.getParentNode().removeChild(userFieldNode);
                return;
            }
            
            // TODO LO/AOO StandardFormatHint.TABLE_KEEP
            // TODO LO/AOO StandardFormatHint.TABLE_REMOVE
        }
        
        if (null == dataValue) {
            // Wenn es ein ValueInterceptor war, dann darf der NIE null zurück geben. Pfff...
            throw new DocumentPlaceholderMissingException(String.format(
                    "Invalid return value (null!) from ValueInterceptor while processing DataValue Key: %s",
                    fieldName));
                    
        }
        
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
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // FORMATIERUNGEN UND STYLE-MANAGEMENT
    ////////////////////////////////////////////////////////////////////////////////////////////////

    private Element prepareStyleElementFor(Element anyStyledElement) {
        final Element style = lookupStyleElementFor(anyStyledElement);
        
        final Element styleCopy = (Element) style.cloneNode(true);
        final String attrStyleName = styleCopy.getAttribute("style:name");
        final String newStyleName = attrStyleName + ".MZ";
        
        styleCopy.setAttribute("style:name", newStyleName);
        styleCopy.removeAttribute("style:display-name");
        
        style.getParentNode().insertBefore(styleCopy, style);
        
        return styleCopy;
    }
    
    private void markStyleTextOrParagraphHiddenFor(Element anyElement) {
        final Element elStyleStyle = prepareStyleElementFor(anyElement);
        final Element styleTextProperties = elementByTagName("style:text-properties", elStyleStyle).orElse(null);
        
        final String attrTextDisplay = styleTextProperties.getAttribute("text:display");
        if (null != attrTextDisplay && attrTextDisplay.isEmpty() == false) {
            return;
        }
        
        styleTextProperties.setAttribute("text:display", "none");
        anyElement.setAttribute("text:style-name", elStyleStyle.getAttribute("style:name"));
    }
    
    private Element lookupStyleElementFor(Element anyElement) {
        final String styleName = lookupStyleNameFor(anyElement);
        assert null != styleName : "null == styleName";
        final NodeList styleStyleList = anyElement.getOwnerDocument()
                .getElementsByTagName("style:style");
        
        for (int i = 0; i < styleStyleList.getLength(); i++) {
            final Element styleStyle = (Element) styleStyleList.item(i);
            final String attrStyleName = styleStyle.getAttribute("style:name");
            
            if (styleName.equals(attrStyleName)) {
                return styleStyle;
            }
        }
        
        return null;
    }
    
    private String lookupStyleNameFor(Element anyElement) {
        // Weg 1:   Es handelt sich um die Standardfälle:
        //          text:style-name
        //          table:style-name
        final String attrTextStyleName = anyElement.getAttribute("text:style-name");
        if (null != attrTextStyleName && attrTextStyleName.isEmpty() == false) {
            return attrTextStyleName;
        }
        
        final String attrTableStyleName = anyElement.getAttribute("table:style-name");
        if (null != attrTableStyleName && attrTableStyleName.isEmpty() == false) {
            return attrTableStyleName;
        }
        
        // Weg 2:   Wir suchen nach einem Attribut mit dem Muster *:style-name
        final NamedNodeMap attrMap = anyElement.getAttributes();
        for (int i = 0; i < attrMap.getLength(); i++) {
            final String attrName = attrMap.item(i).getNodeName();
            
            if (attrName.endsWith(":style-name")) {
                return attrMap.item(i).getNodeValue();
            }
        }
        
        return null;
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // ROUTINEN ZUR INTEGRATION VON GRAFIKEN/BILDERN IN DOKUMENTEN
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private void replaceDrawFrame(Element drawFrame, DataValueMap values) {
        final boolean allNeededElementsExist = elementByTagName("draw:image", drawFrame).isPresent();
        
        if (allNeededElementsExist == false) return;
        
        final ImageValue imageValue;
        {
            final String attrDrawName = drawFrame.getAttribute("draw:name");
            final Optional<ImageValue> valByDrawName = getImageValueByKeyName(attrDrawName, values);
            
            if (valByDrawName.isPresent()) {
                imageValue = valByDrawName.get();
            } else {
                final Element svgTitle = elementByTagName("svg:title", drawFrame).orElse(null);
                if (null == svgTitle) return;
                
                final Optional<ImageValue> valByTitle = getImageValueByKeyName(
                        svgTitle.getTextContent(), values);
                
                if (valByTitle.isPresent()) {
                    imageValue = valByTitle.get();
                } else {
                    return;
                }
            }
        }
        
        setupDrawFrameElement(drawFrame, imageValue);
    }
    
    private void replaceUserFieldWithDrawFrame(Node userFieldNode, ImageValue imageValue) {
        final Element newDrawFrame = createDrawFrameElement(userFieldNode.getOwnerDocument());
        final Element oldUserField = (Element) userFieldNode;
        
        oldUserField.getParentNode().replaceChild(newDrawFrame, oldUserField);
        
        setupDrawFrameElement(newDrawFrame, imageValue);
    }
    
    private Optional<ImageValue> getImageValueByKeyName(String keyName, DataValueMap values) {
        if (null == keyName || keyName.isEmpty() || keyName.trim().isEmpty()) {
            return Optional.empty();
        }
        
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

    
    private void setupDrawFrameElement(Element drawFrame, ImageValue imageValue) {
        final ImageResource imageResource = imageValue.getImageResource();
        
        final String imagePath = registerImageResource(imageResource);
        
        final String imageId = Integer.toString(imageCounter++);
        final String attrDrawName = "Image " + imageId;
        final String attrDrawStyleName = "GrStId" + imageId;
        
        drawFrame.setAttribute("draw:name", attrDrawName);
        
        if (drawFrame.getAttribute("draw:style-name").isEmpty()) {
            // Style-ID vergeben und registrieren
            drawFrame.setAttribute("draw:style-name", attrDrawStyleName);
            
            final NodeList nodeList = drawFrame.getOwnerDocument()
                    .getElementsByTagName("office:automatic-styles");
            
            if (nodeList.getLength() == 0) {
                throw new IllegalStateException("Format error - no office:automatic-styles");
            }
            
            final Element officeAutomaticStyles = (Element) nodeList.item(0);
            officeAutomaticStyles.appendChild(createGraphicsStyleElement(
                    drawFrame.getOwnerDocument(), attrDrawStyleName));
        }
        
        overwriteDrawFrameLink(drawFrame, imagePath);
        applyNonVisualProperties(drawFrame, imageValue);
        applyVisualProperties(drawFrame, imageValue);
    }
    
    private void overwriteDrawFrameLink(Element drawFrame, String imagePath) {
        final Element drawImage = elementByTagName("draw:image", drawFrame).orElse(null);
        drawImage.setAttribute("xlink:href", imagePath);
    }
    
    private void applyNonVisualProperties(Element drawFrame, ImageValue imageValue) {
        final Element svgTitle = elementByTagName("svg:title", drawFrame).orElse(null);
        final Element svgDesc = elementByTagName("svg:desc", drawFrame).orElse(null);
        
        final String orTextSvgTitle = (null == svgTitle) ? "" : svgTitle.getTextContent();
        final String orTextSvgDesc = (null == svgDesc) ? "" : svgDesc.getTextContent();
        
        final String nwTextSvgTitle = imageValue.getTitle().orElse(orTextSvgTitle);
        final String nwTextSvgDesc = imageValue.getDescription().orElse(orTextSvgDesc);
        
        if (null != svgTitle) drawFrame.removeChild(svgTitle);
        if (null != svgDesc) drawFrame.removeChild(svgDesc);
        
        if (null != nwTextSvgTitle && nwTextSvgTitle.isEmpty() == false) {
            final Element nwSvgTitle = drawFrame.getOwnerDocument().createElement("svg:title");
            nwSvgTitle.setTextContent(nwTextSvgTitle);
            drawFrame.appendChild(nwSvgTitle);
        }
        
        if (null != nwTextSvgDesc && nwTextSvgDesc.isEmpty() == false) {
            final Element nwSvgDesc = drawFrame.getOwnerDocument().createElement("svg:desc");
            nwSvgDesc.setTextContent(nwTextSvgDesc);
            drawFrame.appendChild(nwSvgDesc);
        }
    }
    
    private void applyVisualProperties(Element drawFrame, ImageValue imageValue) {
        final String attrSvgHeight = drawFrame.getAttribute("svg:height");
        final String attrSvgWidth = drawFrame.getAttribute("svg:width");
        
        final boolean dimensionExists = !attrSvgHeight.isEmpty() && !attrSvgWidth.isEmpty();
        
        if (dimensionExists && imageValue.isOverwriteDimension() == false) {
            return;
        }
        
        final String newHeight = formatLength(imageValue.getHeight()) + "mm";
        final String newWidth = formatLength(imageValue.getWidth()) + "mm";
        
        drawFrame.setAttribute("svg:height", newHeight);
        drawFrame.setAttribute("svg:width", newWidth);
    }
    
    private Element createDrawFrameElement(Document document) {
        final Element drawFrame = document.createElement("draw:frame");
        drawFrame.setAttribute("draw:name", "");                // überschreiben
        drawFrame.setAttribute("svg:height", "");               // überschreiben
        drawFrame.setAttribute("svg:width", "");                // überschreiben
        drawFrame.setAttribute("draw:style-name", "");          // überschreiben
        drawFrame.setAttribute("style:rel-height", "scale");
        drawFrame.setAttribute("style:rel-width", "scale");
        drawFrame.setAttribute("text:anchor-type", "as-char");
        
        final Element drawImage = (Element) drawFrame.appendChild(document.createElement("draw:image"));
        drawImage.setAttribute("xlink:actuate", "onLoad");
        drawImage.setAttribute("xlink:show", "embed");
        drawImage.setAttribute("xlink:type", "simple");
        drawImage.setAttribute("xlink:href", "Pictures/12345.png"); // überschreiben
        
        return drawFrame;
    }
    
    private Element createGraphicsStyleElement(Document document, String styleId) {
        final Element styleStyle = document.createElement("style:style");
        styleStyle.setAttribute("style:name", styleId);
        styleStyle.setAttribute("style:family", "graphic");
        styleStyle.setAttribute("style:parent-style-name", "Graphics");
        
        final Element styleGraphicProperties = (Element) styleStyle.appendChild(
                document.createElement("style:graphic-properties"));
        
        styleGraphicProperties.setAttribute("fo:background-color", "transparent");
        styleGraphicProperties.setAttribute("fo:border", "none");
        
        return styleStyle;
    }
    
    private String registerImageResource(ImageResource imageResource) {
        final ImageResourceType mimeType;
        {
            final ImageResourceType originMimeType = Objects.requireNonNull(
                    imageResource.getImageFormatType(),
                    "ImageResource#getImageFormatType() == null");
            
            if (originMimeType instanceof OpenDocumentImageResourceType) {
                mimeType = originMimeType;
            } else {
                mimeType = convert(originMimeType, OpenDocumentImageResourceType.values());
            }
        }
        
        if (cacheImageResources.containsKey(imageResource)) {
            return cacheImageResources.get(imageResource);
        }
        
        final boolean isExternalResource = imageResource instanceof LocalImageResource
                || imageResource instanceof ExternalImageResource;
        
        if (isExternalResource && loadAndEmbedExternalImages() == false) {
            // Resource ist extern, und soll anhand der Einstellungen nicht eingebunden werden.
            // Der zurück gegebene Pfad ist dementsprechend extern und wird nicht registriert
            if (imageResource instanceof ExternalImageResource) {
                final String resourceURL =  ((ExternalImageResource) imageResource)
                        .getResourceURL().toString();
                
                this.cacheImageResources.put(imageResource, resourceURL);
                
                return resourceURL;
            } else if (imageResource instanceof LocalImageResource) {
                final String localFilePath = "file:///" + ((LocalImageResource) imageResource)
                        .getLocalResource().toAbsolutePath().toString().replace('\\', '/');
                
                this.cacheImageResources.put(imageResource, localFilePath);
                
                return localFilePath;
            }
        }
        
        // MIME-Type mit internem Dateipfad eintragen
        final String imagePath = "Pictures/img"
                + (UUID.randomUUID().toString().replace("-", "") + ".")
                + mimeType.getFileNameExtensions()[0];
        
        final NodeList manifestNodeList = getDocumentManifest().getElementsByTagName("manifest:manifest");
        if (manifestNodeList.getLength() == 0) {
            throw new IllegalStateException("Internal - no manifest root element found");
        }
        
        final Element manifest = (Element) manifestNodeList.item(0);
        final Element manifestFileEntry = createManifestFileEntryElement(
                Objects.requireNonNull(
                        mimeType.getMimeType(),
                        "ImageResourceType#getMimeType() == null"),
                imagePath);
        
        manifest.appendChild(manifestFileEntry);
        
        // Zum Pfad die Bild-Resource einbinden
        getNewDocumentFile().createNewFileInZip(imagePath);
        getNewDocumentFile().overwrite(imagePath, Objects.requireNonNull(
                imageResource.loadImageData(), "ImageResource#loadData() == null"));
        
        this.cacheImageResources.put(imageResource, imagePath);
        
        return imagePath;
    }
    
    private Element createManifestFileEntryElement(String mimeType, String resourcePath) {
        final Element manifestFileEntry = getDocumentManifest().createElement("manifest:file-entry");
        manifestFileEntry.setAttribute("manifest:media-type", mimeType);
        manifestFileEntry.setAttribute("manifest:full-path", resourcePath);
        return manifestFileEntry;
    }
    
    private String formatLength(double length) {
        return BigDecimal.valueOf(length)
                .setScale(4, RoundingMode.CEILING)
                .toPlainString();
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // XML Manipulations Methoden
    ////////////////////////////////////////////////////////////////////////////////////////////////

    private Node findDocumentBody(Node rootNode) {
        return findNodeByName(rootNode, "office:body");
    }
    
    private Node findUserFieldDeclaration(Node rootNode) {
        return findNodeByName(rootNode, "text:user-field-decls");
    }
    
    private String getFieldName(Node fieldNode) {
        final Element unknownElement = (Element) fieldNode;
        final String elementName = unknownElement.getNodeName();
        
        if (elementName.equals("text:user-field-get")) {
            return getAttribute(fieldNode, "text:name");
        }
        
        if (elementName.equals("text:placeholder")) {
            final String rawFieldName = getTextContent(fieldNode);
            return rawFieldName.replace("<", "").replace(">", "");
        }
        
        throw new IllegalStateException("(Internal) Unknown field type: " + elementName);
    }
    
    private boolean isBodyElement(Node anyNode) {
        if (anyNode instanceof Element == false) return false;
        return "office:body".equals(anyNode.getNodeName());
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
    
    private GenericNodeIterator walkDrawFrames(Node rootNode) {
        return new GenericNodeIterator(rootNode, "draw:frame")
                .noRecursionByElements("table:table");
    }
    
    private GenericNodeIterator walkPlaceholders(Node rootNode) {
        return new GenericNodeIterator(rootNode, "text:placeholder")
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
