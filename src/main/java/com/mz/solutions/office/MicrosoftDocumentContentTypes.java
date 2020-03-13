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

import static com.mz.solutions.office.resources.AbstractOfficeXmlDocumentKeys.INVALID_DOC_FORMAT;
import static com.mz.solutions.office.resources.MessageResources.formatMessage;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 *
 * @author  Riebe, Moritz   (moritz.riebe@mz-entwickler.de)
 */
final class MicrosoftDocumentContentTypes {
    
    private static final String MIME_TYPE_MAIN_DOCUMENT =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    
    private static final String MIME_TYPE_STYLES = 
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    
    private static final String MIME_TYPE_HEADER = 
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    
    private static final String MIME_TYPE_FOOTER = 
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
    
    private static final String MIME_TYPE_FOOTNOTES = 
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
    
    private static final String MIME_TYPE_ENDNOTES = 
            "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private static final String[] EMPTY_STRING_ARRAY = new String[0];
    
    protected static final String ZIP_CONTENT_TYPES = "[Content_Types].xml";
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private final Document docContentTypes;
    
    public MicrosoftDocumentContentTypes(Document docContentTypes) {
        this.docContentTypes = Objects.requireNonNull(docContentTypes, "docContentTypes");
    }
    
    public String getPathForMainDocument() {
        return assertNonEmptyPartArray(findByType(MIME_TYPE_MAIN_DOCUMENT))[0];
    }
    
    public String getPathForMainDocumentRelations() {
        return "word/_rels/document.xml.rels";
    }
    
    public String getPathForMainStyles() {
        return assertNonEmptyPartArray(findByType(MIME_TYPE_STYLES))[0];
    }
    
    public String[] getPathsForHeaders() {
        return findByType(MIME_TYPE_HEADER);
    }
    
    public String[] getPathsForFooters() {
        return findByType(MIME_TYPE_FOOTER);
    }
    
    public String[] getPathsForFootnotes() {
        return findByType(MIME_TYPE_FOOTNOTES);
    }
    
    public String[] getPathsForEndnotes() {
        return findByType(MIME_TYPE_ENDNOTES);
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private Document getDocument() {
        return this.docContentTypes;
    }
    
    private String[] findByType(String mimeType) {
        final NodeList nodeList = getDocument().getElementsByTagName("Types");
        
        if (nodeList.getLength() == 0) {
            return EMPTY_STRING_ARRAY;
        }
        
        final Element elTypes = (Element) nodeList.item(0);
        final NodeList childElements = elTypes.getChildNodes();
        
        final List<String> resultList = new ArrayList<>(3);
        
        for (int elementIndex = 0; elementIndex < childElements.getLength(); elementIndex++) {
            final Node anyChildNode = childElements.item(elementIndex);
            
            if (anyChildNode.getNodeType() != Node.ELEMENT_NODE) {
                continue;
            }
            
            final Element elOverride = (Element) anyChildNode;
            
            if ("Override".equals(elOverride.getNodeName()) == false) {
                continue;
            }
            
            final String attrContentType = elOverride.getAttribute("ContentType");
            
            if (mimeType.equals(attrContentType)) {
                final String attrPartName = elOverride.getAttribute("PartName");
                
                if (attrPartName.isEmpty()) {
                    continue;
                }
                
                resultList.add(convertZipFilePath(attrPartName));
            }
        }
        
        return resultList.isEmpty() ? EMPTY_STRING_ARRAY : resultList.toArray(String[]::new);
    }
    
    private String[] assertNonEmptyPartArray(String[] anyStringArray) {
        if (null == anyStringArray || anyStringArray.length == 0) {
            throw new OfficeDocumentException.InvalidDocumentFormatForImplementation(
                    formatMessage(INVALID_DOC_FORMAT));
        }
        
        return anyStringArray;
    }
    
    private String convertZipFilePath(String path) {
        Objects.requireNonNull(path, "path");

        // Der Pfad in der [Content_Types].xml ist immer mit führemdem Slash angegeben.
        // Wenn dieser vorhanden ist, entfernen wir den führenden Slash.
        
        return path.startsWith("/") ? path.substring(1) : path;
    }
    
}
