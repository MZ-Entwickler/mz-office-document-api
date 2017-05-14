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
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    public ImageResourceType getImageFormatType();
    
    public byte[] loadImageData();
    
}
