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

import com.mz.solutions.office.instruction.DocumentInterceptionContext;
import com.mz.solutions.office.instruction.DocumentInterceptor;
import com.mz.solutions.office.instruction.DocumentInterceptorType;
import com.mz.solutions.office.model.DataValueMap;
import org.w3c.dom.Document;
import org.w3c.dom.Node;

import java.util.List;

class BaseDocumentInterceptorContext extends BaseInstructionContext
        implements DocumentInterceptionContext {

    private DocumentInterceptor interceptor;
    private Document partDocument;
    private Node bodyNode;
    private List<DataValueMap<?>> documentValues;

    public BaseDocumentInterceptorContext(OfficeDocument document) {
        super(document);
    }

    public void setXmlFields(Document partDocument, Node bodyNode) {
        this.partDocument = partDocument;
        this.bodyNode = bodyNode;
    }

    @Override
    public DocumentInterceptor getInterceptor() {
        return interceptor;
    }

    public void setInterceptor(DocumentInterceptor interceptor) {
        this.interceptor = interceptor;
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////

    @Override
    public String getPartName() {
        return getInterceptor().getPartName();
    }

    @Override
    public boolean isXmlBasedDocumentPart() {
        return true;
    }

    @Override
    public DocumentInterceptorType getInterceptorType() {
        return getInterceptor().getInterceptorType();
    }

    @Override
    public List<DataValueMap<?>> getDocumentValues() {
        return documentValues;
    }

    public void setDocumentValues(List<DataValueMap<?>> documentValues) {
        this.documentValues = documentValues;
    }

    @Override
    public List<DataValueMap<?>> getInterceptorValues() {
        return getInterceptor().getInterceptorValues();
    }

    @Override
    public Document getXmlDocument() {
        return partDocument;
    }

    @Override
    public Node getXmlBodyNode() {
        return bodyNode;
    }

}
