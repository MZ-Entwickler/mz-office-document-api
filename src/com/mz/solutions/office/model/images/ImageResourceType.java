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
package com.mz.solutions.office.model.images;

import com.mz.solutions.office.MicrosoftImageResourceType;
import com.mz.solutions.office.OpenDocumentImageResourceType;

/**
 * Zu implementierende Schnittstelle um der jeweiligen Office-Implementierung das Bild-Format
 * anzugeben mit MIME-Type, möglichen Dateinamens-Erweiterungen und einer UI-tauglichen Bezeichnung.
 * 
 * <p>Dazu liegen bereits folgende Vorimplementierungen vor:</p>
 * 
 * <pre>
 *  {@link StandardImageResourceType}       Standard-Formate (BMP, JPG, GIF, PNG)
 *  {@link OpenDocumentImageResourceType}   Unterstützte Formate für Libre-/Open-Office
 *  {@link MicrosoftImageResourceType}      Unterstützte Formate für Microsoft Office
 * </pre>
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
public interface ImageResourceType {
    
    /**
     * Rückgabe einer - optionalen - Beschreibung des Dateiformates.
     * 
     * @return  z.B. {@code 'Windows Bitmap'}.
     */
    public String getFileTypeName();
    
    /**
     * Rückgabe des Mime-types.
     * 
     * @return  z.B. {@code 'image/bmp', 'image/jped'}.
     */
    public String getMimeType();
    
    /**
     * Rückgabe der zutreffenden Datei-Namens-Erweiterungen ohne Punkt und ohne Sternchen.
     * 
     * <p>Die erste Namens-Erweiterung in der Rückgabe ist die Standarderweiterung. Wird z.B. ein
     * Bild mit {@code 'jpe'} übergeben, so kann dieses ggf. umbenannt werden nach {@code 'jpg'}
     * um der Office-Implementierung zu entsprechen.</p>
     * 
     * @return  z.B. {@code ['jpg', 'jpeg', 'jpe', 'jfif']}.
     */
    public String[] getFileNameExtensions();
    
}
