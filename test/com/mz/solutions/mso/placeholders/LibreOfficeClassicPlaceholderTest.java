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
package com.mz.solutions.mso.placeholders;

import com.mz.solutions.office.OfficeDocument;
import com.mz.solutions.office.OfficeDocumentFactory;
import com.mz.solutions.office.result.ResultFactory;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.junit.Test;

public final class LibreOfficeClassicPlaceholderTest extends AbstractClassPlaceholderTest {
    
    @Test
    public void testFile_LibreOffice_PlaceholdersAndUserDefiniedFields_odt() {
        final Path docFile = Paths.get("LibreOffice_PlaceholdersAndUserDefiniedFields.odt");
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newOpenOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(ROOT_IN.resolve(docFile));
        
        document.generate(createDataPage(), ResultFactory.toFile(
                ROOT_OUT.resolve("LibreOffice_PlaceholdersAndUserDefiniedFields_Output.odt")));
    }
    
}
