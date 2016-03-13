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
package com.mz.solutions.office.model;

import org.junit.Ignore;
import org.junit.Test;

/**
 * Oberklasse für DataValue Test: es wird explizit eine Englische und eine Deutsche Version getestet!
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
@Ignore
public class DataValueTest {
    
    @Test(expected = IllegalArgumentException.class)
    public void testConstructorFail_KeyTooShort() {
        new DataValue("N", "Nr. 17");
    }
    
    @Test(expected = IllegalArgumentException.class)
    public void testConstructorFail_KeyContainsWhiteSpaces() {
        new DataValue("KUNDE VORNAME", "Bob");
    }
    
    @Test(expected = IllegalArgumentException.class)
    public void testConstructorFail_TooManyOptions() {
        new DataValue("K_VORNAME", "Huurrz", ValueOptions.KEEP_LINEBREAK, ValueOptions.TABULATOR_TO_FOUR_SPACES, ValueOptions.TABULATOR_TO_WHITESPACE);
    }

}
