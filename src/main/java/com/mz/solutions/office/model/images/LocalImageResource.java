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

import java.nio.file.Path;

/**
 * Bild dessen Quelle extern auf dem/einem lokalen Datenträger ist und welches, soweit möglich und
 * vom Format unterstützt, nicht in das Dokument fest eingebunden werden soll, sondern von der
 * externen Datei bezogen werden soll.
 *
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 * @see ImageResource
 */
public interface LocalImageResource extends ImageResource {

    /**
     * Erzeugt eine lokale Bild-Resource.
     *
     * <p>Identisch zur Methode: {@link ImageResource#useLocalFile(Path, ImageResourceType)}.</p>
     *
     * @param imageFile  Datei
     * @param formatType Bild-Format
     * @return Lokale Bild-Resource
     */
    public static LocalImageResource useLocalFile(Path imageFile, ImageResourceType formatType) {
        return ImageResource.useLocalFile(imageFile, formatType);
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////

    /**
     * Rückgabe des Ortes der lokalen Bild-Resource.
     *
     * <p>Wird von der Office-Implementierung keine lokale Resource unterstützt, wird konventionell
     * weiterhiin {@link #loadImageData()} aufgerufen, welches ggf. die Bild-Datei lädt.</p>
     *
     * @return Pfad zur Bild-Resource
     */
    public Path getLocalResource();

}
