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

import com.mz.solutions.office.resources.MessageResources;

import java.io.IOException;
import java.nio.file.Path;

import static com.mz.solutions.office.resources.MicrosoftDocumentFactoryKeys.NOT_ACCESSIBLE;
import static java.nio.charset.StandardCharsets.US_ASCII;

final class MicrosoftDocumentFactory extends OfficeDocumentFactory {

    private static final byte[] B_ZIP_PK;
    private static final byte[] B_CONTENT_TYPES;
    private static final byte[] B_DOCUMENT;
    private static final byte[] B_WORD_STYLES;
    private static final byte[] B_WORD_FONT_TABLE;
    private static final byte[] B_DOC_PROPS;

    static {
        B_ZIP_PK = "PK".getBytes(US_ASCII);
        B_CONTENT_TYPES = "[Content_Types].xml".getBytes(US_ASCII);
        B_DOCUMENT = "document.xml".getBytes(US_ASCII);
        B_WORD_STYLES = "word/styles".getBytes(US_ASCII);
        B_WORD_FONT_TABLE = "word/fontTable.xml".getBytes(US_ASCII);
        B_DOC_PROPS = "docProps/app.xml".getBytes(US_ASCII);
    }

    protected MicrosoftDocumentFactory() {
        setProperty(MicrosoftProperty.INS_HARD_PAGE_BREAKS, Boolean.TRUE);
        setProperty(MicrosoftProperty.USE_DRAWING_OVER_VML, Boolean.TRUE);
    }

    @Override
    public OfficeDocument openDocument(Path document) {
        if (isFileAccessible(document) == false) {
            throw new IllegalArgumentException(MessageResources.formatMessage(NOT_ACCESSIBLE,
                    /* {0} */ document.toString()));
        }

        return new MicrosoftDocument(this, document);
    }

    @Override
    protected boolean isMyDocumentType(Path document) {
        try {
            return isMyDocumentType0(document);

        } catch (Exception ex) {
            ex.printStackTrace(System.err);
            return false;
        }
    }

    private boolean isMyDocumentType0(Path document)
            throws IOException {

        if (isFileAccessible(document) == false) {
            return false;
        }

        final int MAX_BYTES = (1024 * 1024) * 16; /* max. 16 MB byte */
        final byte[] data = readDataFromFile(document, MAX_BYTES);

        return contains(data, B_ZIP_PK) && (
                contains(data, B_CONTENT_TYPES)
                        || contains(data, B_DOCUMENT)
                        || contains(data, B_DOC_PROPS)
                        || contains(data, B_WORD_FONT_TABLE)
                        || contains(data, B_WORD_STYLES));
    }

}
