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

import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.nio.file.Paths;

public final class LibreOfficeClassicPlaceholderTest extends AbstractClassPlaceholderTest {

    @Test
    public void placeholdersAndUserDefinedFields() {

        String documentPath = AbstractClassPlaceholderTest.class.getResource("LibreOffice_PlaceholdersAndUserDefiniedFields.odt").getPath();
        Path inputPath = Paths.get(documentPath);
        Path outputPath = outputPathOf(inputPath);

        processOpenDocument(createDataPage(), inputPath, outputPath);
    }

}
