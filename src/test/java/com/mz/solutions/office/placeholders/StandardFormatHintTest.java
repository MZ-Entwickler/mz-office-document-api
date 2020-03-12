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

import com.mz.solutions.office.model.hints.*;
import com.mz.solutions.office.AbstractOfficeTest;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataValue;
import org.junit.jupiter.api.Test;

public final class StandardFormatHintTest extends AbstractOfficeTest {
    
    @Test
    public void testFile_StandardFormatHint_LibreOffice_odt() {
        processOpenDocument(
                createDataPage(), "StandardFormatHint_LibreOffice.odt",
                "StandardFormatHint_LibreOffice_Output.odt");
    }
    
    @Test
    public void testFile_StandardFormatHint_MicrosoftOffice_docx() {
        processWordDocument(
                createDataPage(), "StandardFormatHint_MicrosoftWord.docx",
                "StandardFormatHint_MicrosoftWord_Output.docx");
    }
    
    private DataPage createDataPage() {
        final DataPage page = new DataPage();
        
        page.addValue(new DataValue("ANY_VALUE", "Wert ersetzt!"));
        page.addValue(new DataValue("STANDARD_FORMAT_HINT_1", StandardFormatHint.PARAGRAPH_KEEP));
        page.addValue(new DataValue("STANDARD_FORMAT_HINT_2", StandardFormatHint.PARAGRAPH_HIDDEN));
        page.addValue(new DataValue("STANDARD_FORMAT_HINT_3", StandardFormatHint.PARAGRAPH_REMOVE));
        
        page.addValue(new DataValue("FORMAT_HINT_TABLE_KEEP", StandardFormatHint.TABLE_KEEP));
        page.addValue(new DataValue("FORMAT_HINT_TABLE_REMOVE", StandardFormatHint.TABLE_REMOVE));
        
        return page;
    }
    
}
