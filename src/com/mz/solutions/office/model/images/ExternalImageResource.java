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

import java.net.URL;

/**
 * Bild dessen Quelle eine URL ist und, soweit möglich und vom Format unterstützt, nicht direkt ins
 * Dokument eingebunden werden soll, sondern von der URL beim Öffnen des Dokumentes bezogen
 * werden soll.
 * 
 * @see ImageResource
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-entwickler.de)
 */
public interface ExternalImageResource extends ImageResource {
    
    /**
     * Erzeugt eine externe Bild-Resource.
     * 
     * <p>Identisch zur Methode {@link ImageResource#useExternalFile(URL, ImageResourceType)}.</p>
     * 
     * @param imageURL  URL der Bild-Resource
     * @param imgType   Bild-Format
     * 
     * @return          Externe Bild-Resource
     */
    public static ExternalImageResource useExternalFile(URL imageURL, ImageResourceType imgType) {
        return ImageResource.useExternalFile(imageURL, imgType);
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    /**
     * Gibt die URL zurück an der die Bild-Resource gefunden/geladen werden kann.
     * 
     * <p>Unterstützt die Office-Implementierung externe Resourcen, wird diese Methode aufgerufen
     * um die Quelle zu ermitteln. Bei fehlender Unterstützung wird nur {@link #loadImageData()}
     * aufgerufen, welches die Bild-Resource lädt, und normal im Dokument eingebettet.</p>
     * 
     * @return  URL der Bild-Resource, nie {@code null}
     */
    public URL getResourceURL();
    
}
