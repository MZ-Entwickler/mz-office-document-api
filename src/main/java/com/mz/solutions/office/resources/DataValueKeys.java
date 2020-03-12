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
package mz.solutions.office.resources;

import com.mz.solutions.office.model.DataValue; // JavaDoc

/**
 * Intern - Schlüssel für die lokalisierten Texte für {@link DataValue}.
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-entwickler.de)
 */
public interface DataValueKeys {
    
    static final String KEY_TOO_SHORT = "DataValue_KeyTooShort";
    static final String CONTAINS_SPACES = "DataValue_ContainsWhiteSpaces";
    static final String TOO_MANY_VALUE_OPTIONS = "DataValue_TooManyValueOptions";
    static final String OPTIONS_IN_CONFLICT = "DataValue_OptionsInConflict";
    static final String UNKNOWN_OPTION_STATE = "DataValue_UnknownOptionState";
    
}
