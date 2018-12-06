/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2017,   Moritz Riebe     (moritz.riebe@mz-solutions.de)
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

import java.util.Arrays;
import java.util.zip.ZipEntry;

/**
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
final class ZippedEntry {

    ZipEntry entry;
    byte[] data;

    public ZippedEntry cloneEntry() {
        final ZippedEntry copy = new ZippedEntry();

        if (null != data) {
            copy.data = Arrays.copyOf(data, data.length);
        }

        copy.entry = new ZipEntry(entry);

        return copy;
    }

}
