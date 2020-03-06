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
package com.mz.solutions.mso;

import com.mz.solutions.office.OfficeDocument;
import com.mz.solutions.office.OfficeDocumentFactory;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.result.ResultFactory;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.junit.Ignore;

@Ignore
public abstract class AbstractOfficeTest {
    
    protected static final Path ROOT_IN = Paths.get("test-templates");
    protected static final Path ROOT_OUT = ROOT_IN.resolve("output");
    
    protected final void processOpenDocument(DataPage page, Path docInput, Path docOutput) {
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newOpenOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(
                ROOT_IN.resolve(docInput));
        
        document.generate(page, ResultFactory.toFile(
                ROOT_OUT.resolve(docOutput)));
    }
    
    protected final void processOpenDocument(DataPage page, String inFileName, String outFileName) {
        processOpenDocument(page, Paths.get(inFileName), Paths.get(outFileName));
    }
    
    protected final void processWordDocument(DataPage page, Path docInput, Path docOutput) {
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(
                ROOT_IN.resolve(docInput));
        
        document.generate(page, ResultFactory.toFile(
                ROOT_OUT.resolve(docOutput)));
    }
    
    protected final void processWordDocument(DataPage page, String inFileName, String outFileName) {
        processWordDocument(page, Paths.get(inFileName), Paths.get(outFileName));
    }
    
}
