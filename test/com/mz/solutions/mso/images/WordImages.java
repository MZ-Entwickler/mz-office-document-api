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
package com.mz.solutions.mso.images;

import com.mz.solutions.office.MicrosoftProperty;
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
import com.mz.solutions.office.model.interceptor.DataValueResult;
import com.mz.solutions.office.model.interceptor.ValueInterceptor;
import com.mz.solutions.office.result.ResultFactory;
import java.nio.file.Path;
import java.nio.file.Paths;
import static java.util.Arrays.asList;
import java.util.List;
import org.junit.Test;

public class WordImages {
    
    private static final Path ROOT_IN = Paths.get("test-templates");
    private static final Path ROOT_OUT = ROOT_IN.resolve("output");
    
    private static final Path IMG_1 = ROOT_IN.resolve("img_result_1.png");
    private static final Path IMG_2 = ROOT_IN.resolve("img_result_2.png");
    private static final Path IMG_3 = ROOT_IN.resolve("img_result_3.png");
    
    private static final Path INPUT_FILE_1 = ROOT_IN.resolve("WordImagesInDocument.docx");
    private static final Path INPUT_FILE_2 = ROOT_IN.resolve("WordImagesInDocument_Tables.docx");
    
    @Test
    public void testFile_WordImagesInDocument() {
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(INPUT_FILE_1);
        
        document.generate(
                asList(createEmbeddedDataPage("Default Settings (all should get auto embedded) - Page 1"),
                        createEmbeddedDataPage("Default Settings (all should get auto embedded) - Page 2")),
                ResultFactory.toFile(ROOT_OUT.resolve("WordImagesDocument_Output_1.docx")));
        
        document.generate(
                asList(createMixedDataPage("Default Settings (all should get auto embedded) - Page 1"),
                        createMixedDataPage("Default Settings (all should get auto embedded) - Page 2")),
                ResultFactory.toFile(ROOT_OUT.resolve("WordImagesDocument_Output_2.docx")));
        
        document.generate(
                asList(createExternalDataPage("Default Settings (all should get auto embedded) - Page 1"),
                        createExternalDataPage("Default Settings (all should get auto embedded) - Page 2")),
                ResultFactory.toFile(ROOT_OUT.resolve("WordImagesDocument_Output_3.docx")));
        
        document.generate(
                createSharedEmbeddedDataPages("sharing ImageResource and ImageValue"),
                ResultFactory.toFile(ROOT_OUT.resolve("WordImagesDocument_Output_4.docx")));
        
        docFactory.setProperty(MicrosoftProperty.USE_DRAWING_OVER_VML, Boolean.FALSE);
        document.generate(
                asList(createMixedDataPage("Use w:pict over w:drawing (auto embedding)) - Page 1"),
                        createMixedDataPage("Use w:pict over w:drawing (auto embedding)) - Page 2")),
                ResultFactory.toFile(ROOT_OUT.resolve("WordImagesDocument_Output_5.docx")));
        
        docFactory.setProperty(OfficeProperty.IMG_LOAD_AND_EMBED_EXTERNAL, Boolean.FALSE);
        document.generate(
                asList(createMixedDataPage("Use w:pict over w:drawing (NO AUTO EMBEDDING)) - Page 1"),
                        createMixedDataPage("Use w:pict over w:drawing (NO AUTO EMBEDDING)) - Page 2")),
                ResultFactory.toFile(ROOT_OUT.resolve("WordImagesDocument_Output_6.docx")));
        
        docFactory.setProperty(MicrosoftProperty.USE_DRAWING_OVER_VML, Boolean.TRUE);
        document.generate(
                asList(createMixedDataPage("use w:drawing over w:pict (NO AUTO EMBEDDING)) - Page 1"),
                        createMixedDataPage("use w:drawing over w:pict (NO AUTO EMBEDDING)) - Page 2")),
                ResultFactory.toFile(ROOT_OUT.resolve("WordImagesDocument_Output_7.docx")));
    }
    
    @Test
    public void testFile_WordImagesInDocument_WithProperties() {
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(INPUT_FILE_1);
        
        document.generate(
                createEmbeddedDataPageWithProperties(),
                ResultFactory.toFile(ROOT_OUT.resolve("WordImagesDocument_Output_WithProps_1_DefaultSettings.docx")));
        
        docFactory.setProperty(MicrosoftProperty.USE_DRAWING_OVER_VML, Boolean.FALSE);
        document.generate(
                createEmbeddedDataPageWithProperties(),
                ResultFactory.toFile(ROOT_OUT.resolve("WordImagesDocument_Output_WithProps_2_UsePictOverDrawing_Embedded.docx")));
        
        docFactory.setProperty(OfficeProperty.IMG_LOAD_AND_EMBED_EXTERNAL, Boolean.FALSE);
        document.generate(
                createEmbeddedDataPageWithProperties(),
                ResultFactory.toFile(ROOT_OUT.resolve("WordImagesDocument_Output_WithProps_3_UsePictOverDrawing_NoAutoEmbedding.docx")));
        
        docFactory.setProperty(MicrosoftProperty.USE_DRAWING_OVER_VML, Boolean.TRUE);
        document.generate(
                createEmbeddedDataPageWithProperties(),
                ResultFactory.toFile(ROOT_OUT.resolve("WordImagesDocument_Output_WithProps_4_UseDrawingOverPict_NoAutoEmbedding.docx")));
    }
    
    @Test
    public void testFile_WordImagesInDocument_Interception() {
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(INPUT_FILE_1);
        
        final DataPage page = new DataPage();
        page.addValue(new DataValue("TEST_DESCR", "Using interceptors and disabled auto embedding"));
        
        page.addValue(new DataValue("IMG_STANDALONE", ValueInterceptor.callFunction(context -> {
            return DataValueResult.useValue(new DataValue("IMG_STANDALONE",
                    new ImageValue(ImageResource.loadImage(IMG_1, StandardImageResourceType.PNG))));
        })));
        
        page.addValue(new DataValue("IMG_INLINE", ValueInterceptor.callFunction(context -> {
            return DataValueResult.useValue(new DataValue("IMG_INLINE",
                    new ImageValue(ImageResource.loadImage(IMG_2, StandardImageResourceType.PNG))));
        })));
        
        page.addValue(new DataValue("IMG_EXTERNAL", ValueInterceptor.callFunction(context -> {
            return DataValueResult.useValue(new DataValue("IMG_EXTERNAL",
                    new ImageValue(ImageResource.useLocalFile(IMG_3, StandardImageResourceType.PNG))));
        })));
        
        document.generate(page, ResultFactory.toFile(ROOT_OUT.resolve(
                "WordImagesDocument_Output_UsingValueInterceptor.docx")));
    }
    
    @Test
    public void testFile_WordImagesInDocuments_Tables() {
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(INPUT_FILE_2);
        
        final DataPage page = new DataPage();
        page.addValue(new DataValue("IMAGE_1", new ImageValue(
                ImageResource.loadImage(IMG_1, StandardImageResourceType.PNG))));
        
        final DataTable table = new DataTable("IMAGE_TABLE");
        table.addTableRow(new DataTableRow(
                new DataValue("ROW_VALUE", "Image Number 1"),
                new DataValue("IMAGE_1", new ImageValue(
                        ImageResource.loadImage(IMG_1, StandardImageResourceType.PNG)))));
        table.addTableRow(new DataTableRow(
                new DataValue("ROW_VALUE", "Image Number 2"),
                new DataValue("IMAGE_1", new ImageValue(
                        ImageResource.loadImage(IMG_2, StandardImageResourceType.PNG)))));
        table.addTableRow(new DataTableRow(
                new DataValue("ROW_VALUE", "Image Number 3"),
                new DataValue("IMAGE_1", new ImageValue(
                        ImageResource.loadImage(IMG_3, StandardImageResourceType.PNG)))));
        
        page.addTable(table);
        
        document.generate(page, ResultFactory.toFile(ROOT_OUT.resolve(
                "WordImagesDocument_Output_UsingTables.docx")));
    }
    
    private DataPage createEmbeddedDataPageWithProperties() {
        final DataPage page = new DataPage();
        
        page.addValue(new DataValue("TEST_DESCR", "with Properties (Title, Descr)"));
        page.addValue(new DataValue("IMG_STANDALONE",
                new ImageValue(ImageResource.loadImage(IMG_1, StandardImageResourceType.PNG))
                        .setTitle("Standalone Image")
                        .setDescription("Description of a standalone image")));
        
        page.addValue(new DataValue("IMG_INLINE",
                new ImageValue(ImageResource.loadImage(IMG_2, StandardImageResourceType.PNG))
                        .setTitle("Inline Image")
                        .setDescription("Description of a inline image")));
        
        page.addValue(new DataValue("IMG_EXTERNAL",
                new ImageValue(ImageResource.useLocalFile(IMG_3, StandardImageResourceType.PNG))
                        .setTitle("External Image")
                        .setDescription("Description of a external image")));
        
        return page;
    }
    
    private DataPage createEmbeddedDataPage(String appendDescription) {
        final DataPage page = new DataPage();
        
        page.addValue(new DataValue("TEST_DESCR", "Embedded " + appendDescription));
        
        page.addValue(new DataValue("IMG_STANDALONE", new ImageValue(
                ImageResource.loadImage(IMG_1, StandardImageResourceType.PNG))));
        
        page.addValue(new DataValue("IMG_INLINE", new ImageValue(
                ImageResource.loadImage(IMG_2, StandardImageResourceType.PNG))));
        
        page.addValue(new DataValue("IMG_EXTERNAL", new ImageValue(
                ImageResource.loadImage(IMG_3, StandardImageResourceType.PNG))));
        
        return page;
    }
    
    private List<DataPage> createSharedEmbeddedDataPages(String appendDescription) {
        final ImageValue imgValue_1 = new ImageValue(
                ImageResource.loadImage(IMG_1, StandardImageResourceType.PNG));
        
        final ImageValue imgValue_2 = new ImageValue(
                ImageResource.loadImage(IMG_2, StandardImageResourceType.PNG));
        
        final ImageValue imgValue_3 = new ImageValue(
                ImageResource.loadImage(IMG_3, StandardImageResourceType.PNG));
        
        final DataPage page1 = new DataPage();
        page1.addValue(new DataValue("TEST_DESCR", "Shared Image Resources " + appendDescription));
        page1.addValue(new DataValue("IMG_STANDALONE", imgValue_1));
        page1.addValue(new DataValue("IMG_INLINE", imgValue_2));
        page1.addValue(new DataValue("IMG_EXTERNAL", imgValue_3));
        
        final DataPage page2 = new DataPage();
        page2.addValue(new DataValue("TEST_DESCR", "Shared Image Resources " + appendDescription));
        page2.addValue(new DataValue("IMG_STANDALONE", imgValue_1));
        page2.addValue(new DataValue("IMG_INLINE", imgValue_2));
        page2.addValue(new DataValue("IMG_EXTERNAL", imgValue_3));
        
        return asList(page1, page2);
    }
    
    private DataPage createMixedDataPage(String appendDescription) {
        final DataPage page = new DataPage();
        
        page.addValue(new DataValue("TEST_DESCR", "Embedded " + appendDescription));
        
        page.addValue(new DataValue("IMG_STANDALONE", new ImageValue(
                ImageResource.loadImage(IMG_1, StandardImageResourceType.PNG))));
        
        page.addValue(new DataValue("IMG_INLINE", new ImageValue(
                ImageResource.loadImage(IMG_2, StandardImageResourceType.PNG))));
        
        page.addValue(new DataValue("IMG_EXTERNAL", new ImageValue(
                ImageResource.useLocalFile(IMG_3, StandardImageResourceType.PNG))));
        
        return page;
    }
    
    private DataPage createExternalDataPage(String appendDescription) {
        final DataPage page = new DataPage();
        
        page.addValue(new DataValue("TEST_DESCR", "External " + appendDescription));
        
        page.addValue(new DataValue("IMG_STANDALONE", new ImageValue(
                ImageResource.useLocalFile(IMG_1, StandardImageResourceType.PNG))));
        
        page.addValue(new DataValue("IMG_INLINE", new ImageValue(
                ImageResource.useLocalFile(IMG_2, StandardImageResourceType.PNG))));
        
        page.addValue(new DataValue("IMG_EXTERNAL", new ImageValue(
                ImageResource.useLocalFile(IMG_3, StandardImageResourceType.PNG))));
        
        return page;
    }
    
}
