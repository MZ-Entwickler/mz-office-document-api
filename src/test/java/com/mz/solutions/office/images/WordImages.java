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
package com.mz.solutions.office.images;

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
import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import static java.util.Arrays.asList;

public class WordImages extends AbstractImageTest {

    private static final Path INPUT_FILE_1 = Paths.get(WordImages.class.getResource("WordImagesInDocument.docx").getPath());
    private static final Path INPUT_FILE_2 = Paths.get(WordImages.class.getResource("WordImagesInDocument_Tables.docx").getPath());

    @Test
    void testFile_WordDummyImages() {

        String input = LibreOfficeImages.class.getResource("WordDummyImages.docx").getPath();
        Path inputPath = Paths.get(input);
        Path outputPath = outputPathOf(inputPath);

        OfficeDocumentFactory.newMicrosoftOfficeInstance()
                .openDocument(inputPath)
                .generate(create16ColorValuesPage(), ResultFactory.toFile(outputPath));
    }

    @Test
    void testFile_WordImageScaling_wDrawing() {

        String input = LibreOfficeImages.class.getResource("WordImageScaling_wDrawing.docx").getPath();
        Path inputPath = Paths.get(input);
        Path outputPath = outputPathOf(inputPath);

        OfficeDocumentFactory.newMicrosoftOfficeInstance()
                .openDocument(inputPath)
                .generate(createImageScalingPage(), ResultFactory.toFile(outputPath));
    }

    @Test
    void testFile_WordImageScaling_wPict() {

        String input = LibreOfficeImages.class.getResource("WordImageScaling_wPict.docx").getPath();
        Path inputPath = Paths.get(input);
        Path outputPath = outputPathOf(inputPath);

        OfficeDocumentFactory.newMicrosoftOfficeInstance()
                .openDocument(inputPath)
                .generate(createImageScalingPage(), ResultFactory.toFile(outputPath));
    }

    @Test
    void testFile_WordImageScaling_Mergefield() {

        String input = LibreOfficeImages.class.getResource("WordImageScaling_Mergefield.docx").getPath();
        Path inputPath = Paths.get(input);

        final OfficeDocument document = OfficeDocumentFactory.newMicrosoftOfficeInstance()
                .openDocument(inputPath);

        document.generate(createImageScalingPage(), ResultFactory.toFile(outputPathOf(inputPath, "UsingWDrawing")));

        document.getRelatedFactory().setProperty(MicrosoftProperty.USE_DRAWING_OVER_VML, Boolean.FALSE);
        document.generate(createImageScalingPage(), ResultFactory.toFile(outputPathOf(inputPath, "UsingWPict")));

    }

    @Test
    void testFile_WordImagesInDocument() {
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(INPUT_FILE_1);

        document.generate(
                asList(createEmbeddedDataPage("Default Settings (all should get auto embedded) - Page 1"),
                        createEmbeddedDataPage("Default Settings (all should get auto embedded) - Page 2")),
                ResultFactory.toFile(outputPathOf(INPUT_FILE_1, "1")));

        document.generate(
                asList(createMixedDataPage("Default Settings (all should get auto embedded) - Page 1"),
                        createMixedDataPage("Default Settings (all should get auto embedded) - Page 2")),
                ResultFactory.toFile(outputPathOf(INPUT_FILE_1, "2")));

        document.generate(
                asList(createExternalDataPage("Default Settings (all should get auto embedded) - Page 1"),
                        createExternalDataPage("Default Settings (all should get auto embedded) - Page 2")),
                ResultFactory.toFile(outputPathOf(INPUT_FILE_1, "3")));

        document.generate(
                createSharedEmbeddedDataPages("sharing ImageResource and ImageValue"),
                ResultFactory.toFile(outputPathOf(INPUT_FILE_1, "4")));

        docFactory.setProperty(MicrosoftProperty.USE_DRAWING_OVER_VML, Boolean.FALSE);
        document.generate(
                asList(createMixedDataPage("Use w:pict over w:drawing (auto embedding)) - Page 1"),
                        createMixedDataPage("Use w:pict over w:drawing (auto embedding)) - Page 2")),
                ResultFactory.toFile(outputPathOf(INPUT_FILE_1, "5")));

        docFactory.setProperty(OfficeProperty.IMG_LOAD_AND_EMBED_EXTERNAL, Boolean.FALSE);
        document.generate(
                asList(createMixedDataPage("Use w:pict over w:drawing (NO AUTO EMBEDDING)) - Page 1"),
                        createMixedDataPage("Use w:pict over w:drawing (NO AUTO EMBEDDING)) - Page 2")),
                ResultFactory.toFile(outputPathOf(INPUT_FILE_1, "6")));

        docFactory.setProperty(MicrosoftProperty.USE_DRAWING_OVER_VML, Boolean.TRUE);
        document.generate(
                asList(createMixedDataPage("use w:drawing over w:pict (NO AUTO EMBEDDING)) - Page 1"),
                        createMixedDataPage("use w:drawing over w:pict (NO AUTO EMBEDDING)) - Page 2")),
                ResultFactory.toFile(outputPathOf(INPUT_FILE_1, "7")));
    }

    @Test
    void testFile_WordImagesInDocument_WithProperties() {
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(INPUT_FILE_1);

        document.generate(
                createEmbeddedDataPageWithProperties(),
                ResultFactory.toFile(outputPathOf(INPUT_FILE_1, "WithProps_DefaultSettings")));

        docFactory.setProperty(MicrosoftProperty.USE_DRAWING_OVER_VML, Boolean.FALSE);
        document.generate(
                createEmbeddedDataPageWithProperties(),
                ResultFactory.toFile(outputPathOf(INPUT_FILE_1, "WithProps_UsePictOverDrawing_Embedded")));

        docFactory.setProperty(OfficeProperty.IMG_LOAD_AND_EMBED_EXTERNAL, Boolean.FALSE);
        document.generate(
                createEmbeddedDataPageWithProperties(),
                ResultFactory.toFile(outputPathOf(INPUT_FILE_1, "WithProps_UsePictOverDrawing_NoAutoEmbedding")));

        docFactory.setProperty(MicrosoftProperty.USE_DRAWING_OVER_VML, Boolean.TRUE);
        document.generate(
                createEmbeddedDataPageWithProperties(),
                ResultFactory.toFile(outputPathOf(INPUT_FILE_1, "WithProps_UseDrawingOverPict_NoAutoEmbedding")));
    }

    @Test
    void testFile_WordImagesInDocument_Interception() {
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

        document.generate(page, ResultFactory.toFile(outputPathOf(INPUT_FILE_1, "UsingValueInterceptor")));
    }

    @Test
    void testFile_WordImagesInDocuments_Tables() {
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

        document.generate(page, ResultFactory.toFile(outputPathOf(INPUT_FILE_2, "UsingTables")));
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
