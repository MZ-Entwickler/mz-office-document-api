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

import com.mz.solutions.mso.AbstractOfficeTest;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.model.images.ImageResource;
import com.mz.solutions.office.model.images.ImageValue;
import com.mz.solutions.office.model.images.UnitOfLength;
import org.junit.Ignore;

@Ignore
public abstract class AbstractClassPlaceholderTest extends AbstractOfficeTest {
    
    protected final DataPage createDataPage() {
        final DataPage page = new DataPage();
        
        page.addValue(new DataValue("VALUE_1", "Value 1 Replaced"));
        page.addValue(new DataValue("VALUE_2", "Value 2 Replaced"));
        page.addValue(new DataValue("VALUE_3", "Value 3 Replaced"));
        
        final ImageResource imageResourceRed = ImageResource.dummyColorImage(0xFF, 0x00, 0x00);
        final ImageResource imageResourceGreen = ImageResource.dummyColorImage(0x00, 0xFF, 0x00);
        final ImageResource imageResourceBlue = ImageResource.dummyColorImage(0x00, 0x00, 0xFF);
        
        final ImageValue imageValueRed = new ImageValue(imageResourceRed)
                .setTitle("Simple Red Image")
                .setDescription("1x1 Pixel red image with RGB (0xFF, 0x00, 0x00)")
                .setDimension(4.0D, 1.5D, UnitOfLength.CENTIMETERS);
        
        final ImageValue imageValueGreen = new ImageValue(imageResourceGreen)
                .setTitle("Simple Green Image")
                .setDescription("1x1 Pixel green image with RGB (0x00, 0xFF, 0x00)")
                .setDimension(4.0D, 1.5D, UnitOfLength.CENTIMETERS);
        
        final ImageValue imageValueBlue = new ImageValue(imageResourceBlue)
                .setTitle("Simple Blue Image")
                .setDescription("1x1 Pixel blue image with RGB (0x00, 0x00, 0xFF)")
                .setDimension(4.0D, 1.5D, UnitOfLength.CENTIMETERS);
        
        page.addValue(new DataValue("IMAGE_RED", imageValueRed));
        page.addValue(new DataValue("IMAGE_GREEN", imageValueGreen));
        page.addValue(new DataValue("IMAGE_BLUE", imageValueBlue));
        
        return page;
    }
    
}
