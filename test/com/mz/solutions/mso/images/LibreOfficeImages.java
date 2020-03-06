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
package com.mz.solutions.mso.images;

import static com.mz.solutions.mso.images.AbstractImageTest.IMG_1;
import com.mz.solutions.office.OfficeDocument;
import com.mz.solutions.office.OfficeDocumentFactory;
import com.mz.solutions.office.OfficeProperty;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataTable;
import com.mz.solutions.office.model.DataTableRow;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.model.images.ImageResource;
import com.mz.solutions.office.model.images.ImageValue;
import com.mz.solutions.office.model.images.StandardImageResourceType;
import com.mz.solutions.office.result.ResultFactory;
import java.nio.file.Path;
import static java.util.Arrays.asList;
import org.junit.Test;

public class LibreOfficeImages extends AbstractImageTest {
    
    private static final Path INPUT_FILE_1 = ROOT_IN.resolve("LibreOfficeImagesInDocument.odt");
    private static final Path INPUT_FILE_2 = ROOT_IN.resolve("OpenDocumentByWord.odt");
    private static final Path INPUT_FILE_3 = ROOT_IN.resolve("LibreOfficeImagesInDocument_Tables.odt");
    
    @Test
    public void testFile_LibreOfficeDummyImages() {
        OfficeDocumentFactory.newOpenOfficeInstance()
                .openDocument(ROOT_IN.resolve("LibreOfficeDummyImages.odt"))
                .generate(create16ColorValuesPage(), ResultFactory
                        .toFile(ROOT_OUT.resolve("LibreOfficeDummyImages_Output.odt")));
    }
    
    @Test
    public void testFile_LibreOfficeImageScaling_Image() {
        final OfficeDocument document = OfficeDocumentFactory.newOpenOfficeInstance()
                .openDocument(ROOT_IN.resolve("LibreOfficeImageScaling_Image.odt"));
        
        document.generate(
                createImageScalingPage(),
                ResultFactory.toFile(ROOT_OUT.resolve("LibreOfficeImageScaling_Image_Output.odt")));
    }
    
    @Test
    public void testFile_LibreOfficeImageScaling_UserDefinedField() {
        final OfficeDocument document = OfficeDocumentFactory.newOpenOfficeInstance()
                .openDocument(ROOT_IN.resolve("LibreOfficeImageScaling_UserDefinedField.odt"));
        
        document.generate(
                createImageScalingPage(),
                ResultFactory.toFile(ROOT_OUT.resolve("LibreOfficeImageScaling_UserDefinedField_Output.odt")));
    }
    
    @Test
    public void testFile_LibreOfficeImagesInDocument() {
        testFile0(INPUT_FILE_1, "byLibreOffice", false);
    }
    
    @Test
    public void testFile_OpenDocumentByWord() {
        testFile0(INPUT_FILE_2, "byWord", false);
    }
    
    @Test
    public void testFile_LibreOfficeImagesInDocument_Extern() {
        testFile0(INPUT_FILE_1, "byLibreOffice_Link", true);
    }
    
    @Test
    public void testFile_OpenDocumentByWord_Extern() {
        testFile0(INPUT_FILE_2, "byWord_Link", true);
    }
    
    @Test
    public void testFile_LibreOfficeImagesInDocument_Tables() {
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newOpenOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(INPUT_FILE_3);
        
        final DataPage page = new DataPage();
        page.addValue(new DataValue("IMG_STANDALONE", new ImageValue(
                ImageResource.loadImage(IMG_1, StandardImageResourceType.PNG))));
        
        final DataTable table = new DataTable("IMAGE_TABLE");
        table.addTableRow(new DataTableRow(
                new DataValue("ROW_VALUE", "Image Number 1"),
                new DataValue("IMG_STANDALONE", new ImageValue(
                        ImageResource.loadImage(IMG_1, StandardImageResourceType.PNG)))));
        table.addTableRow(new DataTableRow(
                new DataValue("ROW_VALUE", "Image Number 2"),
                new DataValue("IMG_STANDALONE", new ImageValue(
                        ImageResource.loadImage(IMG_2, StandardImageResourceType.PNG)))));
        table.addTableRow(new DataTableRow(
                new DataValue("ROW_VALUE", "Image Number 3"),
                new DataValue("IMG_STANDALONE", new ImageValue(
                        ImageResource.loadImage(IMG_3, StandardImageResourceType.PNG)))));
        
        page.addTable(table);
        
        document.generate(page, ResultFactory.toFile(ROOT_OUT.resolve(
                "LibreOfficeImgDoc_WithTable.odt")));
    }

    
    private void testFile0(Path fileName, String namePart, boolean keepExternal) {
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newOpenOfficeInstance();
        if (keepExternal) {
            docFactory.setProperty(OfficeProperty.IMG_LOAD_AND_EMBED_EXTERNAL, Boolean.FALSE);
        }
        
        final OfficeDocument document = docFactory.openDocument(fileName);
        
        document.generate(
                asList(createDataPage(), createDataPage()),
                ResultFactory.toFile(ROOT_OUT.resolve("LibreOfficeImgDoc_" + namePart + ".odt")));
    }
    
    private DataPage createDataPage() {
        final DataPage page = new DataPage();
        
        page.addValue(new DataValue("IMG_STANDALONE", new ImageValue(
                ImageResource.loadImage(IMG_1, StandardImageResourceType.PNG))));
        
        page.addValue(new DataValue("IMG_INLINE", new ImageValue(
                ImageResource.loadImage(IMG_2, StandardImageResourceType.PNG))
                .setTitle("Title of an inline image")
                .setDescription("Description of an inline image")));
        
        page.addValue(new DataValue("IMG_EXTERNAL", new ImageValue(
                ImageResource.useLocalFile(IMG_3, StandardImageResourceType.PNG))));
        
        return page;
    }
    
}
