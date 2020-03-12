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
package com.mz.solutions.office.tables;

import com.mz.solutions.office.AbstractOfficeTest;

import java.nio.file.Path;
import java.nio.file.Paths;

abstract class AbstractTableTest extends AbstractOfficeTest {

    protected static final Path NESTED_TABLES = Paths.get(AbstractTableTest.class.getResource("NestedTables.docx").getPath());
    protected static final Path NORMAL_TABLES_DOCX = Paths.get(AbstractTableTest.class.getResource("NormalTables.docx").getPath());
    protected static final Path NORMAL_TABLES_ODT = Paths.get(AbstractTableTest.class.getResource("NormalTables.odt").getPath());
}
