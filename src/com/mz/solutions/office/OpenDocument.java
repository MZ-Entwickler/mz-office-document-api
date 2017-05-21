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
import com.mz.solutions.office.model.DataMap;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataTable;
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
import java.nio.file.Path;
import java.util.IdentityHashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.UUID;
import static mz.solutions.office.resources.AbstractOfficeXmlDocumentKeys.INVALID_DOC_FORMAT;
import static mz.solutions.office.resources.MessageResources.formatMessage;
import static mz.solutions.office.resources.OpenDocumentKeys.NO_DATA;
import static mz.solutions.office.resources.OpenDocumentKeys.UNKNOWN_FORMATTING_CHAR;
import static mz.solutions.office.resources.OpenDocumentKeys.UNKNOWN_PLACE_HOLDER;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

final class OpenDocument extends AbstractOfficeXmlDocument {
    
    private static final String ZIP_DOC_CONTENT = "content.xml";
    private static final String ZIP_DOC_STYLES = "styles.xml";
    private static final String ZIP_MANIFEST = "META-INF/manifest.xml";
    
    private final MyInterceptionContext interceptionContext = new MyInterceptionContext();
    
    private final Document sourceManifest;
    
    private volatile Document newManifest;
    private volatile Document newContent;
    private volatile Document newStyles;
    
    private volatile ZIPDocumentFile newDocumentFile;
    
    private volatile int imageCounter = 16_000;
    private volatile Map<ImageResource, String> cacheImageResources = new IdentityHashMap<>();
    
    public OpenDocument(OfficeDocumentFactory factory, Path document) {
        super(factory, document, ZIP_DOC_CONTENT, ZIP_DOC_STYLES);
        
        try {
            this.sourceManifest = loadFileAsXml(ZIP_MANIFEST);
        } catch (Exception errorWhileLoading) {
            throw new OfficeDocumentException.InvalidDocumentFormatForImplementation(
                    formatMessage(INVALID_DOC_FORMAT), errorWhileLoading);
        }
    }

    @Override
    protected String getImplementedOfficeName() {
        return "Apache OpenOffice 4.x / LibreOffice";
    }

    
    @Override
    protected ZIPDocumentFile createAndFillDocument(
            final Iterator<DataPage> dataPages) {
        
        this.newDocumentFile = sourceDocumentFile.cloneDocument();
        this.imageCounter = 16_000;
        this.cacheImageResources.clear();
        
        try {
            fillDocuments0(dataPages);
            return newDocumentFile;
        } finally {
            this.cacheImageResources.clear();
            this.newContent = this.newStyles = this.newManifest = null;
            this.newDocumentFile = null;
        }
    }
    
    private void fillDocuments0(final Iterator<DataPage> dataPages) {
        this.newContent = (Document) sourceContent.cloneNode(true);
        this.newStyles = (Document) sourceStyles.cloneNode(true);
        this.newManifest = (Document) sourceManifest.cloneNode(true);
        
        removeUserFieldDeclaration(newContent);
        
        final Node nodeContentBody = findDocumentBody(newContent);
        
        // Leeren Content-Body um alle Datensätze anfügen zu können
        final Node newFullContentBody = nodeContentBody.cloneNode(true);
        final NodeList bodyChildNodes = newFullContentBody.getChildNodes();
        
        for (int i = 0; i < bodyChildNodes.getLength(); i++) {
            newFullContentBody.removeChild(bodyChildNodes.item(i));
        }
        
        boolean missingDataPages = true;
        
        while (dataPages.hasNext()) {
            final DataPage nextPage = dataPages.next();
            final Node newContentBody = nodeContentBody.cloneNode(true);
            
            replaceDocumentTree(newContentBody, nextPage);
            
            // Alle Elemente im leeren Content-Body anfügen
            final NodeList filledNodes = newContentBody.getChildNodes();
            
            for (int i = 0; i < filledNodes.getLength(); i++) {
                newFullContentBody.appendChild(filledNodes.item(i));
            }
            
            missingDataPages = false;
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

        overwrite(newDocumentFile, ZIP_DOC_CONTENT, newContent);
        overwrite(newDocumentFile, ZIP_DOC_STYLES, newStyles);
        overwrite(newDocumentFile, ZIP_MANIFEST, newManifest);
    }
    
    private void removeUserFieldDeclaration(Node documentRoot) {
        final Node textUserFieldDecls =  findUserFieldDeclaration(documentRoot);
        
        if (null == textUserFieldDecls) {
            return; // Keine Benutzerdefinierten Felder
        }
        
        final Node parentNode = textUserFieldDecls.getParentNode();
        
        parentNode.removeChild(textUserFieldDecls);
    }
    
    private void replaceDocumentTree(Node node, DataMap valueMap) {
        replaceUserFieldsAndImages(node, valueMap);
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
    
    private void replaceUserFieldsAndImages(Node documentBody, DataValueMap value) {
        for (Node drawFrameNode : walkDrawFrames(documentBody)) {
            replaceDrawFrame((Element) drawFrameNode, value);
        }
        
        for (Node userFieldNode : walkUserFields(documentBody)) {
            replaceUserFieldNode(userFieldNode, value);
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
        
        final int dataRowIndex = tableRows.size() == 1 ? 0 : 1;
        final Node tableDataRow = tableRows.get(dataRowIndex);
        
        // Alle Zeilen die NICHT die zu wiederholende Datenzeilen sind
        // ganz normal ersetzen
        for (int i = 0; i < tableRows.size(); i++) {
            if (dataRowIndex == i /* current */) {
                continue;
            }
            
            replaceUserFieldsAndImages(tableRows.get(i), tableData.get());
        }
        
        final Iterator<DataTableRow> rowIterator = tableData.get().iterator();
        while (rowIterator.hasNext()) {
            final DataTableRow rowData = rowIterator.next();
            final Node newTableRow = tableDataRow.cloneNode(true);
            
            replaceDocumentTree(newTableRow, rowData);
            
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
        if (dataValue.isExtendedValue() && dataValue.extendedValue() instanceof ImageValue) {
            // Ohhh Platzhalter ist ein Bild!
            replaceUserFieldWithDrawFrame(userFieldNode, (ImageValue) dataValue.extendedValue());
            return;
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
        final Element newDrawFrame = createDrawFrameElement();
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
            
            final NodeList nodeList = newContent.getElementsByTagName("office:automatic-styles");
            if (nodeList.getLength() == 0) {
                throw new IllegalStateException("Format error - no office:automatic-styles");
            }
            
            final Element officeAutomaticStyles = (Element) nodeList.item(0);
            officeAutomaticStyles.appendChild(createGraphicsStyleElement(attrDrawStyleName));
        }
        
        overwriteDrawFrameLink(drawFrame, imagePath);
        applyNonVisualProperties(drawFrame, imageValue);
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
            final Element nwSvgTitle = newContent.createElement("svg:title");
            nwSvgTitle.setTextContent(nwTextSvgTitle);
            drawFrame.appendChild(nwSvgTitle);
        }
        
        if (null != nwTextSvgDesc && nwTextSvgDesc.isEmpty() == false) {
            final Element nwSvgDesc = newContent.createElement("svg:desc");
            nwSvgDesc.setTextContent(nwTextSvgDesc);
            drawFrame.appendChild(nwSvgDesc);
        }
    }
    
    private Element createDrawFrameElement() {
        final Element drawFrame = newContent.createElement("draw:frame");
        drawFrame.setAttribute("draw:name", "");                // überschreiben
        drawFrame.setAttribute("svg:height", "1.5cm");          // überschreiben
        drawFrame.setAttribute("svg:width", "4cm");             // überschreiben
        drawFrame.setAttribute("draw:style-name", "");          // überschreiben
        drawFrame.setAttribute("style:rel-height", "scale");
        drawFrame.setAttribute("style:rel-width", "scale");
        drawFrame.setAttribute("text:anchor-type", "as-char");
        
        final Element drawImage = (Element) drawFrame.appendChild(newContent.createElement("draw:image"));
        drawImage.setAttribute("xlink:actuate", "onLoad");
        drawImage.setAttribute("xlink:show", "embed");
        drawImage.setAttribute("xlink:type", "simple");
        drawImage.setAttribute("xlink:href", "Pictures/12345.png"); // überschreiben
        
        return drawFrame;
    }
    
    private Element createGraphicsStyleElement(String styleId) {
        final Element styleStyle = newContent.createElement("style:style");
        styleStyle.setAttribute("style:name", styleId);
        styleStyle.setAttribute("style:family", "graphic");
        styleStyle.setAttribute("style:parent-style-name", "Graphics");
        
        final Element styleGraphicProperties = (Element) styleStyle.appendChild(
                newContent.createElement("style:graphic-properties"));
        
        styleGraphicProperties.setAttribute("fo:background-color", "transparent");
        styleGraphicProperties.setAttribute("fo:border", "none");
        
        return styleStyle;
    }
    
    private String registerImageResource(ImageResource imageResource) {
        final ImageResourceType mimeType = Objects.requireNonNull(
                imageResource.getImageFormatType(),
                "ImageResource#getImageFormatType() == null");
        
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
        
        final NodeList manifestNodeList = newManifest.getElementsByTagName("manifest:manifest");
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
        newDocumentFile.createNewFileInZip(imagePath);
        newDocumentFile.overwrite(imagePath, Objects.requireNonNull(
                imageResource.loadImageData(), "ImageResource#loadData() == null"));
        
        this.cacheImageResources.put(imageResource, imagePath);
        
        return imagePath;
    }
    
    private Element createManifestFileEntryElement(String mimeType, String resourcePath) {
        final Element manifestFileEntry = newManifest.createElement("manifest:file-entry");
        manifestFileEntry.setAttribute("manifest:media-type", mimeType);
        manifestFileEntry.setAttribute("manifest:full-path", resourcePath);
        return manifestFileEntry;
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
    
    private GenericNodeIterator walkDrawFrames(Node rootNode) {
        return new GenericNodeIterator(rootNode, "draw:frame")
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
