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
package com.mz.solutions.mso.placeholders;

import com.mz.solutions.office.OfficeDocument;
import com.mz.solutions.office.OfficeDocumentFactory;
import com.mz.solutions.office.result.ResultFactory;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.junit.Test;

public final class MicrosoftWordClassicPlaceholderTest extends AbstractClassPlaceholderTest {
    
    @Test
    public void testFile_Word_Placeholders_docx() {
        final Path docFile = Paths.get("Word_Placeholders.docx");
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(ROOT_IN.resolve(docFile));
        
        document.generate(createDataPage(), ResultFactory.toFile(
                ROOT_OUT.resolve("Word_Placeholders_Output.docx")));
    }
    
}
