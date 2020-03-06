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
package com.mz.solutions.office.model.images;

/**
 * Standard Bild-Formate die von beiden Office-Implementierungen unterstützt werden.
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-entwickler.de)
 */
public enum StandardImageResourceType implements ImageResourceType {

    BMP ("Windows Bitmap", "image/bmp", array("bmp", "dib", "rle")),
    JPG ("JPEG File Interchange Format", "image/jpeg", array("jpg", "jpeg", "jfif", "jpe", "jif")),
    GIF ("Graphics Interchange Format", "image/gif", array("gif")),
    PNG ("Portable Network Graphics", "image/png", array("png"));
    
    private static String[] array(String ... elements) { return elements; }
        
    private final String description;
    private final String mimeType;
    
    private final String[] fileNameExtensions;

    private StandardImageResourceType(String description, String mimeType, String[] fileNameExtensions) {
        this.description = description;
        this.mimeType = mimeType;
        this.fileNameExtensions = fileNameExtensions;
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
