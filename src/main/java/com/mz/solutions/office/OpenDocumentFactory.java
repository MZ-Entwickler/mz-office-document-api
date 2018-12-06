/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2018,   Moritz Riebe     (moritz.riebe@mz-solutions.de),
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

import java.io.IOException;
import java.nio.file.Path;

import static com.mz.solutions.office.resources.MessageResources.formatMessage;
import static com.mz.solutions.office.resources.OpenDocumentFactoryKeys.NOT_ACCESSIBLE;
import static java.nio.charset.StandardCharsets.US_ASCII;

/**
 * Factory Implementierung für OpenOffice.
 *
 * @author Riebe, Moritz       (moritz.riebe@mz-solutions.de)
 * @date 2015-03-11
 */
final class OpenDocumentFactory extends OfficeDocumentFactory {

    private static final byte[] B_ZIP_PK;
    private static final byte[] B_MIME;
    private static final byte[] B_OASIS;
    private static final byte[] B_META;
    private static final byte[] B_MANIFEST;
    private static final byte[] B_CONTENT;

    static {
        B_ZIP_PK = "PK".getBytes(US_ASCII);
        B_MIME = "application/vnd.oasis.opendocument.text".getBytes(US_ASCII);
        B_OASIS = "urn:oasis:names:tc:opendocument:xmlns".getBytes(US_ASCII);
        B_META = "office:document-meta".getBytes(US_ASCII);
        B_MANIFEST = "META-INF/manifest.xml".getBytes(US_ASCII);
        B_CONTENT = "content.xml".getBytes(US_ASCII);
    }

    @Override
    public OfficeDocument openDocument(Path document) {
        if (isFileAccessible(document) == false) {
            throw new IllegalArgumentException(formatMessage(NOT_ACCESSIBLE,
                    /* {0} */ document.toString()));
        }

        return new OpenDocument(this, document);
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

        final int MAX_BYTES = (1024 * 1024) * 16; /* max. 16 MB Bytes */
        final byte[] data = readDataFromFile(document, MAX_BYTES);

        return contains(data, B_ZIP_PK) && (
                contains(data, B_MIME)
                        || contains(data, B_OASIS)
                        || contains(data, B_META)
                        || contains(data, B_MANIFEST)
                        || contains(data, B_CONTENT));
    }

}
