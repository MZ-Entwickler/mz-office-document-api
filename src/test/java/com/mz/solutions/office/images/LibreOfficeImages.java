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
package com.mz.solutions.office.images;

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
import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.nio.file.Paths;

import static java.util.Arrays.asList;

public class LibreOfficeImages extends AbstractImageTest {

    private static final Path INPUT_FILE_1 = Paths.get(LibreOfficeImages.class.getResource("LibreOfficeImagesInDocument.odt").getPath());
    private static final Path INPUT_FILE_2 = Paths.get(LibreOfficeImages.class.getResource("OpenDocumentByWord.odt").getPath());
    private static final Path INPUT_FILE_3 = Paths.get(LibreOfficeImages.class.getResource("LibreOfficeImagesInDocument_Tables.odt").getPath());

    @Test
    public void testFile_LibreOfficeDummyImages() {

        String input = LibreOfficeImages.class.getResource("LibreOfficeDummyImages.odt").getPath();
        Path inputPath = Paths.get(input);
        Path outputPath = outputPathOf(inputPath);

        OfficeDocumentFactory.newOpenOfficeInstance()
                .openDocument(inputPath)
                .generate(create16ColorValuesPage(), ResultFactory.toFile(outputPath));
    }

    @Test
    public void testFile_LibreOfficeImageScaling_Image() {

        String input = LibreOfficeImages.class.getResource("LibreOfficeImageScaling_Image.odt").getPath();
        Path inputPath = Paths.get(input);
        Path outputPath = outputPathOf(inputPath);

        OfficeDocumentFactory.newOpenOfficeInstance()
                .openDocument(inputPath)
                .generate(createImageScalingPage(), ResultFactory.toFile(outputPath));
    }

    @Test
    public void testFile_LibreOfficeImageScaling_UserDefinedField() {

        String input = LibreOfficeImages.class.getResource("LibreOfficeImageScaling_UserDefinedField.odt").getPath();
        Path inputPath = Paths.get(input);
        Path outputPath = outputPathOf(inputPath);

        OfficeDocumentFactory.newOpenOfficeInstance()
                .openDocument(inputPath)
                .generate(createImageScalingPage(), ResultFactory.toFile(outputPath));
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

        document.generate(page, ResultFactory.toFile(outputPathOf(INPUT_FILE_3, "WithTable")));
    }


    private void testFile0(Path fileName, String namePart, boolean keepExternal) {

        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newOpenOfficeInstance();
        if (keepExternal) {
            docFactory.setProperty(OfficeProperty.IMG_LOAD_AND_EMBED_EXTERNAL, Boolean.FALSE);
        }

        docFactory.openDocument(fileName)
                .generate(asList(createDataPage(), createDataPage()), ResultFactory.toFile(outputPathOf(fileName, namePart)));
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
