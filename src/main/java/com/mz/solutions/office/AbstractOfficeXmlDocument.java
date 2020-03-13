/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2020,   Moritz Riebe     (moritz.riebe@mz-entwickler.de)
 *                       Andreas Zaschka  (andreas.zaschka@mz-entwickler.de)
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

import com.mz.solutions.office.OfficeDocumentException.FailedDocumentGenerationException;
import com.mz.solutions.office.OfficeDocumentException.InvalidDocumentFormatForImplementation;
import com.mz.solutions.office.instruction.DocumentInterceptor;
import com.mz.solutions.office.instruction.DocumentInterceptorType;
import com.mz.solutions.office.instruction.DocumentProcessingInstruction;
import com.mz.solutions.office.instruction.HeaderFooterInstruction;
import com.mz.solutions.office.model.DataMap;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.model.DataValueMap;
import com.mz.solutions.office.model.images.ImageResourceType;
import com.mz.solutions.office.model.interceptor.InterceptionContext;
import static com.mz.solutions.office.resources.AbstractOfficeXmlDocumentKeys.CREATION_FAILED;
import static com.mz.solutions.office.resources.AbstractOfficeXmlDocumentKeys.IMPL_NAME_ERR;
import static com.mz.solutions.office.resources.AbstractOfficeXmlDocumentKeys.INVALID_DOC_FORMAT;
import static com.mz.solutions.office.resources.MessageResources.formatMessage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import static java.util.Collections.disjoint;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.Predicate;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

abstract class AbstractOfficeXmlDocument extends OfficeDocument {

    /**
     * Enthält die Referenz zur einen Factory-Implementierung.
     */
    protected final OfficeDocumentFactory myOfficeFactory;
    
    // Originale Daten - Sollten NICHT verändert werden!!
    private final ZIPDocumentFile sourceDocumentFile;
    
    // XML Factories
    private final DocumentBuilderFactory docBuilderFactory;
    private final DocumentBuilder docBuilder;
    
    {
        docBuilderFactory = DocumentBuilderFactory.newInstance();
        docBuilderFactory.setIgnoringComments(true);
        
        try {
            docBuilder = docBuilderFactory.newDocumentBuilder();
        } catch (ParserConfigurationException ex) {
            throw new IllegalStateException(ex.getMessage(), ex);
        }
    }
    
    private final TransformerFactory transformerFactory;
    private final Transformer transformer;
    
    {
        transformerFactory = TransformerFactory.newInstance();
        
        try {
            transformer = transformerFactory.newTransformer();
        } catch (TransformerConfigurationException ex) {
            throw new IllegalStateException(ex.getMessage(), ex);
        }
    }
    
    private DocumentProcessingInstruction[] myInstructions = NO_INSTRUCTIONS;
    private volatile BaseDocumentInterceptorContext documentInterceptorContext;
    private volatile BaseHeaderFooterContext headerFooterContext;

    private volatile ZIPDocumentFile newDocumentFile;
    private volatile Map<String, Document> documentParts;
    
    public AbstractOfficeXmlDocument(OfficeDocumentFactory myFactory, Path document) {
        this.myOfficeFactory = myFactory;
        
        try {
            this.sourceDocumentFile = new ZIPDocumentFile(document);
        } catch (Exception errorWhileLoading) {
            throw new InvalidDocumentFormatForImplementation(
                    formatMessage(INVALID_DOC_FORMAT), errorWhileLoading);
        }
    }

    @Override
    public OfficeDocumentFactory getRelatedFactory() {
        return myOfficeFactory;
    }
    
    private Document loadFileAsXml(String zipItemFilename) {
        return bytesToXml(sourceDocumentFile.read(zipItemFilename));
    }
    
    protected final Document bytesToXml(byte[] bytes) {
        final InputStream inFile = new ByteArrayInputStream(bytes);
        
        try {
            return docBuilder.parse(inFile);
            
        } catch (SAXException | IOException ex) {
            throw new OfficeDocumentException
                    .InvalidDocumentFormatForImplementation(ex);
        }
    }
    
    protected final ZIPDocumentFile getNewDocumentFile() {
        return newDocumentFile;
    }
    
    protected final ZIPDocumentFile getSourceDocumentFile() {
        return sourceDocumentFile;
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // DOCUMENT PARTS
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    protected final Document getDocumentPart(String partName) {
        if (documentParts.containsKey(partName)) {
            return documentParts.get(partName);
        }
        
        final Document partDocument = loadFileAsXml(partName);
        documentParts.put(partName, partDocument);
        
        return partDocument;
    }
    
    private void writeChangedDocumentParts() {
        for (String partName : documentParts.keySet()) {
            final Document partDocument = documentParts.get(partName);
            overwrite(partName, partDocument);
        }
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // DOCUMENT INTERCEPTORS
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    protected abstract String normalizePartName(String inputPartName);
    
    private void handleDocumentInterceptors(DocumentInterceptorType interceptorType, List<DataValueMap<?>> documentValues) {
        for (DocumentProcessingInstruction anyInstruction : listInstructions()) {
            if (anyInstruction instanceof DocumentInterceptor == false) continue;
            
            final DocumentInterceptor interceptor = (DocumentInterceptor) anyInstruction;
            
            if (interceptorType.equals(interceptor.getInterceptorType())) {
                runDocumentInterceptor(interceptor, documentValues);
            }
        }
    }
    
    private void runDocumentInterceptor(DocumentInterceptor interceptor, List<DataValueMap<?>> documentValues) {
        final String partName = normalizePartName(interceptor.getPartName());
        final Document partDocument = getDocumentPart(partName);
        final Element bodyNode = partDocument.getDocumentElement();
        
        this.documentInterceptorContext.setInterceptor(interceptor);
        this.documentInterceptorContext.setXmlFields(partDocument, bodyNode);
        this.documentInterceptorContext.setDocumentValues(documentValues);
        
        interceptor.callInterceptor(documentInterceptorContext);
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private String getSafeImplementationName() {
        final String name = getImplementedOfficeName();
        
        if (null == name || name.trim().isEmpty()) {
            throw new IllegalStateException(formatMessage(IMPL_NAME_ERR,
                    /* {0} */ getClass().getSimpleName()));
        }
        
        return name;
    }
    
    protected abstract String getImplementedOfficeName();
    
    protected final void overwrite(String filename, Node domRootNode) {
        getNewDocumentFile().overwrite(filename, xmlToBytes(domRootNode));
    }
    
    protected final byte[] xmlToBytes(Node domRootNode) {
        final int bufSize = 32 * 1024;
        final ByteArrayOutputStream outXML = new ByteArrayOutputStream(bufSize);
        
        domRootNode.normalize();
        
        try {
            transformer.transform(
                    new DOMSource(domRootNode),
                    new StreamResult(outXML));
            
        } catch (TransformerException ex) {
            throw new FailedDocumentGenerationException(
                    formatMessage(CREATION_FAILED), ex);
        }
        
        return outXML.toByteArray();
    }
    
    @Override
    protected byte[] generate(
            final Iterator<DataPage> dataPages,
            final DocumentProcessingInstruction ... instructions)
            throws IOException, OfficeDocumentException
    {
        final ByteArrayOutputStream byteOut = new ByteArrayOutputStream(128 * 1024);

        this.newDocumentFile = sourceDocumentFile.cloneDocument();
        
        this.documentInterceptorContext = new BaseDocumentInterceptorContext(this);
        this.headerFooterContext = new BaseHeaderFooterContext(this);
        
        this.documentParts = new HashMap<>();
        
        final List<DataValueMap<?>> documentValues = toDataValueMap(dataPages);

        try {
            this.myInstructions = checkInstructions(instructions);
            
            handleDocumentInterceptors(DocumentInterceptorType.BEFORE_GENERATION, documentValues);
            
            createAndFillDocument((Iterator) documentValues.iterator());
            
            handleDocumentInterceptors(DocumentInterceptorType.AFTER_GENERATION, documentValues);
            
            writeChangedDocumentParts();
            
            this.newDocumentFile.writeTo(byteOut);
        } finally {
            this.myInstructions = NO_INSTRUCTIONS;
            this.documentInterceptorContext = null;
            this.headerFooterContext = null;
            this.documentParts = null;
            this.newDocumentFile = null;
        }
        
        return byteOut.toByteArray();
    }
    
    private List<DataValueMap<?>> toDataValueMap(Iterator<DataPage> dataPageIterator) {
        final List<DataValueMap<?>> resultList = new LinkedList<>();
        dataPageIterator.forEachRemaining(resultList::add);
        return resultList;
    }
    
    private DocumentProcessingInstruction[] checkInstructions(
            DocumentProcessingInstruction[] instructions)
    {
        if (null == instructions) return NO_INSTRUCTIONS;
        if (instructions.length == 0) return instructions;
        
        for (int instrIndex = 0; instrIndex < instructions.length; instrIndex++) {
            if (null == instructions[instrIndex]) {
                throw new NullPointerException("instructions[@index = " + instrIndex + "]");
            }
        }
        
        return instructions;
    }

    protected final DocumentProcessingInstruction[] listInstructions() {
        return this.myInstructions;
    }
    
    protected final boolean hasInstructions() {
        return listInstructions().length > 0;
    }
    
    protected abstract void createAndFillDocument(Iterator<DataPage> dataPages);
    
    /**
     * Nimmt den übergebenen MIME-Type an und versucht zu diesem einen passenden zu finden aus der
     * eigenen Liste der MIME-Typen.
     * 
     * <p>Ist der MIME-Type direkt in der Liste {@code mimeTypeList}, wird dieser zurück gegeben.
     * Ansonsten wird zuerst anhand des MIME-Types (String) vergleichen und gesucht und danach nach
     * einer passenden Datennamens-Erweiterung. Wurde kein passender MIME-Type aus
     * {@code mimeTypeList} gefunden, wird der ursprünglich übergebene zurück gegeben.</p>
     * 
     * @param mimeType      Im Daten-Modell übergebener MIME-Type.
     * @param mimeTypeList  Liste der von der Office-Implementierung unterstützten MIME-Types.
     * @return              Soweit wie möglich passender MIME-Type zur Implementierung, ansonsten
     *                      der ursprünglich in {@code mimeType} übergebene.
     */
    protected final ImageResourceType convert(
            final ImageResourceType mimeType, final ImageResourceType[] mimeTypeList)
    {
        final String inMimeType = mimeType.getMimeType();
        
        // Schauen ob wir den selben haben anhand des MIME-Types direkt
        for (ImageResourceType singleType : mimeTypeList) {
            if (inMimeType.equalsIgnoreCase(singleType.getMimeType())) {
                return singleType;
            }
        }
        
        // Wohl ein Schuss im Ofen, also suchen wir nach einer Dateinamens-Erweiterung die passen
        // könnte.
        final String[] inExtensions = mimeType.getFileNameExtensions();
        final List<String> inExtList = Arrays.asList(inExtensions);

        for (ImageResourceType singleType : mimeTypeList) {
            final List<String> otherExtList = Arrays.asList(singleType.getFileNameExtensions());
            
            if (disjoint(inExtList, otherExtList) == false) {
                return singleType;
            }
        }
        
        return mimeType;
    }
    
    protected final boolean hasHeaderFooterInstructions() {
        for (DocumentProcessingInstruction anyInstruction : listInstructions()) {
            if (anyInstruction instanceof HeaderFooterInstruction) return true;
        }
        return false;
    }
    
    protected Optional<DataMap<?>> callHeaderFooterInstruction(String name, boolean header) {
        name = (null == name ? "" : name);
        
        if (header) {
            headerFooterContext.setupAsHeader(name);
        } else {
            headerFooterContext.setupAsFooter(name);
        }
        
        for (DocumentProcessingInstruction anyInstruction : listInstructions()) {
            if (anyInstruction instanceof HeaderFooterInstruction == false) continue;
            
            final HeaderFooterInstruction instruction = (HeaderFooterInstruction) anyInstruction;
            final Optional<DataMap<?>> result = instruction.processHeaderFooter(headerFooterContext);
            
            if (result.isPresent()) {
                return result;
            }
        }
        
        return Optional.empty();
    }

    ////////////////////////////////////////////////////////////////////////////
    // Methoden zum Abfragen der Standard-Einstellungen
    ////////////////////////////////////////////////////////////////////////////
    
    /**
     * Überprüfung ob fehlende Werte/Platzhalter ignoriert werden.
     * 
     * @return      sollen ignoriert werden wenn {@code true}
     */
    protected boolean ignoreMissingValues() {
        final Boolean errOnValueMissing = myOfficeFactory
                .getProperty(OfficeProperty.ERR_ON_MISSING_VAL);
        
        return Boolean.FALSE.equals(errOnValueMissing);
    }
    
    /**
     * Überprüfung ob ein Versionskonflikt ignoriert werden soll.
     * 
     * @return      soll ignoriert werden wenn {@code true}
     */
    protected boolean ignoreVersionMismatch() {
        final Boolean errOnVersionMismatch = myOfficeFactory
                .getProperty(OfficeProperty.ERR_ON_VER_MISMATCH);
        
        return Boolean.FALSE.equals(errOnVersionMismatch);
    }
    
    /**
     * Überprüft ob bei fehlenden DataPages die Vorgang leise beendet werden
     * soll und kein Dokument angelegt werden soll.
     * 
     * @return      bei {@code true} Methode leise beenden und kein Dokument
     *              anlegen
     */
    protected boolean ignoreMissingDataPages() {
        final Boolean errOnNoData = myOfficeFactory
                .getProperty(OfficeProperty.ERR_ON_NO_DATA);
        
        return Boolean.FALSE.equals(errOnNoData);
    }
    
    /**
     * Überprüft ob Bild-Resourcen von externen Quellen geladen und direkt ins Dokument eingefügt
     * werden sollen.
     * 
     * @return      {@code true}, bedeutet Bild-Resource laden (in RAM) und direkt in das Dokument
     *              schreiben (einbetten) ohne die externe Quelle weiterzureichen.
     */
    protected boolean loadAndEmbedExternalImages() {
        final Boolean imgLoadAndEmbed = myOfficeFactory.getProperty(
                OfficeProperty.IMG_LOAD_AND_EMBED_EXTERNAL);
        
        return Boolean.TRUE.equals(imgLoadAndEmbed);
    }
    
    ////////////////////////////////////////////////////////////////////////////
    // UMGANG MIT INTERCEPTOR-VALUES (CALLBACK MECHANISMUS)
    ////////////////////////////////////////////////////////////////////////////
    
    protected DataValue handleInterception(DataValue value, InterceptionContext context) {
        return ValueInterceptorUtil.callInterceptors(value, context);
    }
    
    ////////////////////////////////////////////////////////////////////////////
    // Generische XML Helfer Methoden
    ////////////////////////////////////////////////////////////////////////////
    
    protected Node findNodeByName(Node node, String elementName) {
        if (node.getNodeType() == Node.ELEMENT_NODE) {
            if (node.getNodeName().equals(elementName)) {
                return node;
            }
        }
        
        NodeList childNodes = node.getChildNodes();
        for (int i = 0; i < childNodes.getLength(); i++) {
            Node result = findNodeByName(childNodes.item(i), elementName);
            
            if (null != result) {
                return result;
            }
        }
        
        return null;
    }
    
    protected Optional<Element> elementByTagName(String elementName, Node elementNode) {
        if (elementNode instanceof Element == false) {
            throw new IllegalStateException("internal error anyNode != type Element");
        }
        
        final NodeList nodeList = ((Element) elementNode).getElementsByTagName(elementName);
        
        for (int nodeIndex = 0; nodeIndex < nodeList.getLength(); nodeIndex++) {
            final Node anyNode = nodeList.item(nodeIndex);
            
            if (anyNode instanceof Element) {
                return Optional.of((Element) anyNode);
            }
        }
        return Optional.empty();
    }
    
    protected List<Node> nodesByTagName(String elementName, Node elementNode) {
        if (elementNode instanceof Element == false) {
            throw new IllegalStateException("internal error anyNode != type Element");
        }
        
        return toFlatNodeList(((Element) elementNode).getElementsByTagName(elementName));
    }
    
    protected String getAttribute(Node elementNode, String attributeName) {
        if (elementNode instanceof Element) {
            final String attrValue =  ((Element) elementNode).getAttribute(attributeName);
            return (null == attrValue) ? "" : attrValue;
        }
        
        NamedNodeMap attrMap = elementNode.getAttributes();
        
        if (null == attrMap) {
            return "";
        }
        
        Node textNode = attrMap.getNamedItem(attributeName);
        
        if (null == textNode) {
            return "";
        }
        
        final String resultText = textNode.getNodeValue();
        
        return null == resultText ? "" : resultText;
    }
    
    protected String getTextContent(Node elementNode) {
        if (elementNode.hasChildNodes() == false) {
            return "";
        }
        
        final List<Node> childs = toFlatNodeList(elementNode.getChildNodes());
        
        if (childs.size() == 1) { // nur einer ist ja easy
            if (childs.get(0).getNodeType() == Node.TEXT_NODE) {
                final String text = childs.get(0).getNodeValue();
                return null == text ? "" : text;
            } else {
                return ""; // irgendwas anderes
            }
        }
        
        final StringBuilder textBuilder = new StringBuilder();
        for (Node anyChild : childs) {
            if (anyChild.getNodeType() != Node.TEXT_NODE) {
                continue;
            }
            
            textBuilder.append(anyChild.getNodeValue());
        }
        
        return textBuilder.toString();
    }
    
    protected List<Node> toFlatNodeList(NodeList nodeList) {
        final List<Node> result = new ArrayList<>(nodeList.getLength());
        
        for (int i = 0; i < nodeList.getLength(); i++) {
            result.add(nodeList.item(i));
        }
        
        return result;
    }
    
    protected Stream<Node> streamNodes(Node rootNode) {
        return StreamSupport.stream(
                new StreamIterator(rootNode).spliterator(), false);
    }
    
    protected Stream<Node> streamNodes(
            Node rootNode, Predicate<Node> stopFunction) {
        
        return StreamSupport.stream(new StreamIterator(rootNode, stopFunction)
                .spliterator(), false);
    }
    
    protected Stream<Node> streamElements(Node rootNode) {
        return streamNodes(rootNode).filter(this::filterElements);
    }
    
    protected Stream<Node> streamElements(Node rootNode,
            Predicate<Node> stopFunction) {
        
        return streamNodes(rootNode, stopFunction).filter(this::filterElements);
    }
    
    protected boolean filterElements(Node node) {
        return node.getNodeType() == Node.ELEMENT_NODE;
    }
    
    protected boolean excludeNode(Node node, Node otherNode) {
        return node != otherNode;
    }
    
    protected boolean hasAttribute(Node node, String name) {
        return getAttribute(node, name).trim().isEmpty() == false;
    }
    
    public boolean isElement(Node node, String name) {
        return node.getNodeType() == Node.ELEMENT_NODE
                && node.getNodeName().equals(name);
    }
    
    public boolean containsElement(Node node, String elementName) {
        return streamElements(node)
                .filter(n -> isElement(node, elementName))
                .findFirst()
                .isPresent();
    }
    
    protected void removeNode(Node node) {
        node.getParentNode().removeChild(node);
    }
    
    ////////////////////////////////////////////////////////////////////////////
    
    /**
     * Implementierung eines Iterators über einzelne Nodes anhand eines
     * vordefinierten Filters.
     */
    protected final static class GenericNodeIterator implements Iterable<Node> {
        private final List<Node> fieldNodes = new LinkedList<>();
        private final String elementName;
        private String[] stopElements = {};
        private final Node rootNode;
        private boolean skipFirstStopper = false;
            
        public GenericNodeIterator (Node rootNode, String elementName) {
            this.elementName = elementName;
            this.rootNode = rootNode;
        }

        public GenericNodeIterator noRecursionByElements(
                String ... stopElements) {

            this.stopElements = stopElements;

            return this;
        }
        
        public GenericNodeIterator doSkipFirstStopElement() {
            this.skipFirstStopper = true;
            return this;
        }

        private void walkNode(Node node, String elementName, int deep) {
            if (null == node) {
                return;
            }
            
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                if (node.getNodeName().equals(elementName)) {
                    fieldNodes.add(node);
                    return;
                }

                if (!(skipFirstStopper && deep == 0)) {
                    if (null != stopElements && stopElements.length > 0) {
                        for (String element : stopElements) {
                            if (node.getNodeName().equals(element)) {
                                return; // STOP
                            }
                        }
                    }
                }

            }

            NodeList children = node.getChildNodes();

            for (int i = 0; i < children.getLength(); i++) {
                walkNode(children.item(i), elementName, deep + 1);
            }
        }

        public List<Node> asList() {
            fieldNodes.clear();
            walkNode(rootNode, elementName, 0);
            
            return fieldNodes;
        }

        @Override
        public Iterator<Node> iterator() {
            return asList().iterator();
        }
    }
    
    ////////////////////////////////////////////////////////////////////////////
    // Iterator-Implementierung für Streams
    ////////////////////////////////////////////////////////////////////////////
    
    private static class StreamIterator implements Iterable<Node> {
        
        private final List<Node> nodes = new LinkedList<>();
        private final Predicate<Node> stopFunction;
        
        public StreamIterator(Node rootNode, Predicate<Node> stopFunction) {
            this.stopFunction = stopFunction;
            walkingTheTree(rootNode);
        }
        
        public StreamIterator(Node rootNode) {
            this(rootNode, (n) -> false);
        }
        
        private void walkingTheTree(Node node) {
            if (stopFunction.test(node)) {
                return;
            }
            
            if (node.hasChildNodes() == false) {
                nodes.add(node);
                return;
            }
            
            final NodeList childs = node.getChildNodes();
            int length = childs.getLength();
            
            for (int i = 0; i < length; i++) {
                walkingTheTree(childs.item(i));
            }
            
            nodes.add(node);
        }

        @Override
        public Iterator<Node> iterator() {
            return nodes.iterator();
        }
        
    }
    
}
