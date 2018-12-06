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
package com.mz.solutions.office;

import com.mz.solutions.office.model.images.ImageResourceType;
import com.mz.solutions.office.model.images.StandardImageResourceType;

/**
 * Von Microsoft Office unterstütze Bild-Formate.
 *
 * <p>Enthält auch alle Standardformate aus {@link StandardImageResourceType}.</p>
 *
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
public enum MicrosoftImageResourceType implements ImageResourceType {

    BMP("Windows Bitmap", "image/bmp", array("bmp", "dib", "rle")),
    JPG("JPEG File Interchange Format", "image/jpeg", array("jpg", "jpeg", "jfif", "jpe", "jif")),
    GIF("Graphics Interchange Format", "image/gif", array("gif")),
    PNG("Portable Network Graphics", "image/png", array("png")),
    TIF("Tag Image File Format", "image/tiff", array("tif", "tiff")),

    EMF("Windows Enhanced Metafile", "	image/x-emf", array("emf")),
    EMZ("Compressed Windows Enhanced Metafile", "image/x-emf", array("emz")),

    WMF("Windows Metafile", "image/x-wmf", array("wmf")),
    WMZ("Compressed Windows Metafile", "image/x-wmf", array("wmz")),

    PCZ("Compressed Macintosh PICT", "image/x-pict", array("pcz")),
    PCT("Macintosh PICT", "image/x-pict", array("pct", "pict")),

    EPS("Encapsulated PostScript", "image/x-eps", array("eps", "epsf", "epsi")),

    WPF("WordPerfect Graphics", "image/x-wpg", array("wpg"));

    private final String description;
    private final String mimeType;
    private final String[] fileNameExtensions;

    private MicrosoftImageResourceType(
            String description, String mimeType, String[] fileNameExtensions) {
        this.description = description;
        this.mimeType = mimeType;
        this.fileNameExtensions = fileNameExtensions;
    }

    private static String[] array(String... elements) {
        return elements;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getFileTypeName() {
        return description;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getMimeType() {
        return mimeType;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String[] getFileNameExtensions() {
        return fileNameExtensions;
    }

}
