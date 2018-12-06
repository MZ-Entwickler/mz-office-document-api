/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2018,   Moritz Riebe     (moritz.riebe@mz-solutions.de),
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

/**
 * Einsetzen und Ersetzen von Bildern in Vorlagen-Dokumenten.
 *
 * <p>Nutz {@link com.mz.solutions.office.extension.ExtendedValue} um in
 * {@link com.mz.solutions.office.model.DataValue} als Wert Bilder angeben zu können die im
 * Ersetzungsprozess ins Ziel-Dokument eingefügt (an der Stelle des Text-Platzhalters) oder
 * ersetzt (an der Stelle eines bestehendes Bildes als Bild-Platzhalter) werden sollen.</p>
 *
 * <pre>
 *  final ImageResource myImage = ImageResource.loadImage(
 *          Paths.get("myImage.bmp"), StandardImageResourceType.BMP);
 *
 *  final ImageValue imageValue = new ImageValue(myImage)
 *          .setTitle("My Image");
 *
 *  myDataModel.addValue(new DataValue("IMAGE", imageValue));
 * </pre>
 *
 * @see com.mz.solutions.office.model.images.ImageValue    Bild-Wert mit zugeordneter Bild-Resource
 * @see com.mz.solutions.office.model.images.ImageResource Bild-Resource (laden/erzeugen/verlinken)
 */
package com.mz.solutions.office.model.images;
