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
package com.mz.solutions.office.placeholders;

import com.mz.solutions.office.OfficeDocument;
import com.mz.solutions.office.OfficeDocumentFactory;
import com.mz.solutions.office.result.ResultFactory;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.nio.file.Paths;

public final class LibreOfficeClassicPlaceholderTest extends AbstractClassPlaceholderTest {

    @Nested
    class LibreOffice {

        protected final Path docFile = Paths.get(
                AbstractClassPlaceholderTest.class.getResource(
                        "LibreOffice_PlaceholdersAndUserDefiniedFields.odt").getPath());

        @Test
        public void placeholdersAndUserDefinedFields() {

            final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newOpenOfficeInstance();
            final OfficeDocument document = docFactory.openDocument(docFile);

            document.generate(createDataPage(), ResultFactory.toFile(
                    TESTS_OUTPUT_PATH.resolve("LibreOffice_PlaceholdersAndUserDefiniedFields_Output.odt")));
        }
    }

}
