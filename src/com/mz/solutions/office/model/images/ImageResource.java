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
package com.mz.solutions.office.model.images;

import java.io.InputStream;
import java.io.UncheckedIOException;
import java.net.URL;
import java.nio.file.Path;
import java.util.Objects;

/**
 * Eigentliches Bild mit Quelle und/oder Daten des Bildes.
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
public interface ImageResource {
    
    public static ImageResource loadImage(Path imageFile, ImageResourceType formatType)
            throws UncheckedIOException
    {
        Objects.requireNonNull(imageFile, "imageFile");
        Objects.requireNonNull(formatType, "formatType");
        
        return new StandardImageResource.EagerEmbedImgResFile(imageFile, formatType);
    }
    
    public static ImageResource loadImage(InputStream imageStream, ImageResourceType formatType)
            throws UncheckedIOException
    {
        Objects.requireNonNull(imageStream, "imageStream");
        Objects.requireNonNull(formatType, "formatType");
        
        return new StandardImageResource.EagerEmbedImgResData(imageStream, formatType);
    }
    
    public static ImageResource loadImage(byte[] imageData, ImageResourceType formatType) {
        Objects.requireNonNull(imageData, "imageData");
        Objects.requireNonNull(formatType, "formatType");
        
        return new StandardImageResource.EagerEmbedImgResData(imageData, formatType);
    }
    
    public static ImageResource loadImageLazy(Path imageFile, ImageResourceType formatType) {
        Objects.requireNonNull(imageFile, "imageFile");
        Objects.requireNonNull(formatType, "formatType");
        
        return new StandardImageResource.LazyEmbedImgResFile(imageFile, formatType);
    }
    
    public static LocalImageResource useLocalFile(Path imageFile, ImageResourceType imgType) {
        Objects.requireNonNull(imageFile, "imageFile");
        Objects.requireNonNull(imgType, "imgType");
        
        return new StandardImageResource.LocalImageResourceImpl(imageFile, imgType);
    }
    
    public static ExternalImageResource useExternalFile(URL imageURL, ImageResourceType imgType) {
        Objects.requireNonNull(imageURL, "imageURL");
        Objects.requireNonNull(imgType, "imgType");
        
        return new StandardImageResource.ExternalImgResImpl(imageURL, imgType);
    }
    
    public static ImageResource dummyColorImage(int rgbColor) {
        return dummyColorImage(
                (rgbColor >> 16) & 0xFF,
                (rgbColor >>  8) & 0xFF,
                (rgbColor >>  0) & 0xFF);
    }
    
    public static ImageResource dummyColorImage(int red, int green, int blue) {
        red = (red < 0x00 ? 0x00 : (red > 0xFF ? 0xFF : red));
        green = (green < 0x00 ? 0x00 : (green > 0xFF ? 0xFF : green));
        blue = (blue < 0x00 ? 0x00 : (blue > 0xFF ? 0xFF : blue));
        
        final byte[] bmpImageData = new byte[] {
            (byte) 0x42, (byte) 0x4D, (byte) 0x3A, (byte) 0x00, (byte) 0x00, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x36, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x28, (byte) 0x00, (byte) 0x00, (byte) 0x00,
            (byte) 0x01, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x01, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x01, (byte) 0x00, (byte) 0x18, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x04, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0xC4, (byte) 0x0E, (byte) 0x00, (byte) 0x00,
            (byte) 0xC4, (byte) 0x0E, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00,
            (byte) blue, (byte) green, (byte) red, (byte) 0x00
        };
        
        return loadImage(bmpImageData, StandardImageResourceType.BMP);
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    public ImageResourceType getImageFormatType();
    
    public byte[] loadImageData();
    
}
