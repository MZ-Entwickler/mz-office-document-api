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
package com.mz.solutions.office;

import com.mz.solutions.office.model.images.ImageResourceType;
import com.mz.solutions.office.model.images.StandardImageResourceType;

/**
 * Von OpenOffice/LibreOffice unterstützte Bild-Formate.
 * 
 * <p>Enthält auch alle Standardformate aus {@link StandardImageResourceType}.</p>
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-entwickler.de)
 */
public enum OpenDocumentImageResourceType implements ImageResourceType {

    BMP ("Windows Bitmap", "image/bmp", array("bmp", "dib", "rle")),
    JPG ("JPEG File Interchange Format", "image/jpeg", array("jpg", "jpeg", "jfif", "jpe", "jif")),
    GIF ("Graphics Interchange Format", "image/gif", array("gif")),
    PNG ("Portable Network Graphics", "image/png", array("png")),
    TIF ("Tag Image File Format", "image/tiff", array("tif", "tiff")),
    
    EMF ("Windows Enhanced Metafile", "	image/x-emf", array("emf")),
    EMZ ("Compressed Windows Enhanced Metafile", "image/x-emf", array("emz")),
    
    WMF ("Windows Metafile", "image/x-wmf", array("wmf")),
    WMZ ("Compressed Windows Metafile", "image/x-wmf", array("wmz")),
    
    PCZ ("Compressed Macintosh PICT", "image/x-pict", array("pcz")),
    PCT ("Macintosh PICT", "image/x-pict", array("pct", "pict")),
    
    EPS ("Encapsulated PostScript", "image/x-eps", array("eps", "epsf", "epsi")),
    
    DXF ("AutoCAD Interchange Format", "image/vnd.dxf", array("dxf")),
    MET ("OS/2 Metafile", "image/x-met", array("met")),
    MOV ("QuickTime File Format", "video/quicktime", array("mov", "qt")),
    PBM ("Portable Bitmap", "image/x-portable-bitmap", array("pbm", "pgm", "ppm", "pnm")),
    PCX ("Zsoft Paintbrush", "image/vnd.zbrush.pcx", array("pcx")),
    PSD ("Adobe Photoshop", "image/x-psd", array("psd")),
    RAS ("Sun Raster Image", "image/x-ras", array("ras")),
    SGF ("StarWriter Graphics Format", "image/x-sgf", array("sgf")),
    SGV ("StarDraw 2.0", "image/x-sgv", array("sgv")),
    SVG ("Scalable Vector Graphics", "image/svg", array("svg", "svgz")),
    SVM ("StarView Metafile", "image/x-svm", array("svm")),
    TGA ("Truevision Targa", "image/x-tga", array("tga")),
    XBM ("X Bitmap", "image/x-xbm", array("xbm")),
    XPM ("X PixMap", "image/x-xpm", array("xmp"));
    
    private static String[] array(String ... elements) { return elements; }
        
    private final String description;
    private final String mimeType;
    
    private final String[] fileNameExtensions;

    private OpenDocumentImageResourceType(
            String description, String mimeType, String[] fileNameExtensions)
    {
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
