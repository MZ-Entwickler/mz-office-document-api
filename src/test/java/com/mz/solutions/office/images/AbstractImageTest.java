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
package com.mz.solutions.office.images;

import com.mz.solutions.office.AbstractOfficeTest;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataTable;
import com.mz.solutions.office.model.DataTableRow;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.model.images.ImageResource;
import com.mz.solutions.office.model.images.ImageValue;
import com.mz.solutions.office.model.images.UnitOfLength;

import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static com.mz.solutions.office.model.images.StandardImageResourceType.PNG;

abstract class AbstractImageTest extends AbstractOfficeTest {

    protected static final String packageName = "images";
    protected static final Path IMG_1 = TEST_SOURCE_DIRECTORY.resolve(packageName).resolve("img_result_1.png");
    protected static final Path IMG_2 = TEST_SOURCE_DIRECTORY.resolve(packageName).resolve("img_result_2.png");
    protected static final Path IMG_3 = TEST_SOURCE_DIRECTORY.resolve(packageName).resolve("img_result_3.png");

    protected final List<DataPage> createImageScalingPage() {
        final ImageResource image1 = ImageResource.loadImage(IMG_1, PNG);
        final ImageResource image2 = ImageResource.loadImage(IMG_2, PNG);
        final ImageResource image3 = ImageResource.loadImage(IMG_3, PNG);

        final DataPage[] resultPages = new DataPage[2];

        { // PAGE 1 - OVERWRITE IMAGE DIMENSION
            resultPages[0] = new DataPage();
            resultPages[0].addValue(new DataValue("TEST_DESCR", "OverwriteDimension = True"));

            resultPages[0].addValue(new DataValue("IMAGE_1", new ImageValue(image1)
                    .setTitle("Image 1").setDescription("2cm x 2cm")
                    .setDimension(2, 2, UnitOfLength.CENTIMETERS)
                    .setOverwriteDimension(true)));

            resultPages[0].addValue(new DataValue("IMAGE_2", new ImageValue(image2)
                    .setTitle("Image 2").setDescription("4cm x 4cm")
                    .setDimension(4, 4, UnitOfLength.CENTIMETERS)
                    .setOverwriteDimension(true)));

            resultPages[0].addValue(new DataValue("IMAGE_3", new ImageValue(image3)
                    .setTitle("Image 3").setDescription("15.5cm x 1.0cm")
                    .setDimension(15.5, 1.0, UnitOfLength.CENTIMETERS)
                    .setOverwriteDimension(true)));
        }

        { // PAGE 2 - DO NOT OVERWRITE IMAGE DIMENSION
            resultPages[1] = new DataPage();
            resultPages[1].addValue(new DataValue("TEST_DESCR", "OverwriteDimension = False"));

            resultPages[1].addValue(new DataValue("IMAGE_1", new ImageValue(image1)
                    .setTitle("Image 1").setDescription("2cm x 2cm")
                    .setDimension(2, 2, UnitOfLength.CENTIMETERS)
                    .setOverwriteDimension(false)));

            resultPages[1].addValue(new DataValue("IMAGE_2", new ImageValue(image2)
                    .setTitle("Image 2").setDescription("4cm x 4cm")
                    .setDimension(4, 4, UnitOfLength.CENTIMETERS)
                    .setOverwriteDimension(false)));

            resultPages[1].addValue(new DataValue("IMAGE_3", new ImageValue(image3)
                    .setTitle("Image 3").setDescription("15.5cm x 1.0cm")
                    .setDimension(15.5, 1.0, UnitOfLength.CENTIMETERS)
                    .setOverwriteDimension(false)));
        }

        return Arrays.asList(resultPages);
    }

    protected final DataPage create16ColorValuesPage() {
        final DataPage resultPage = new DataPage();
        final DataTable tableColors = new DataTable("TCOLORS");

        final String[] hexValues = {
                "000000", "9D9D9D", "FFFFFF", "BE2633",
                "E06F8B", "493C2B", "A46422", "EB8931",
                "F7E26B", "2F484E", "44891A", "A3CE27",
                "1B2632", "005784", "31A2F2", "B2DCEF"
        };

        for (int i = 0; i < hexValues.length; i++) {
            tableColors.addTableRow(new DataTableRow(
                    colorValue("#" + hexValues[i]),
                    colorImage(Integer.parseInt(hexValues[i], 16))));
        }

        resultPage.addTable(tableColors);

        return resultPage;
    }

    private DataValue colorValue(String value) {
        return new DataValue("COLOR_VALUE", value);
    }

    private DataValue colorImage(int rgbColor) {
        return new DataValue("COLOR_IMAGE", new ImageValue(ImageResource.dummyColorImage(rgbColor))
                .setDimension(12.5, 1, UnitOfLength.CENTIMETERS));
    }

}
