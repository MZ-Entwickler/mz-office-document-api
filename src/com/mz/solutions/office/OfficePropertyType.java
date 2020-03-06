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
package com.mz.solutions.office;

/**
 * Schnittstelle für zulässige Einstellungsoptionen (Schlüssel).
 * 
 * <p>Der verfügbaren Optionen sollten statisch angelegt sein und per
 * Referenz vergbleichbar sein.</p>
 * 
 * @author  Riebe, Moritz       (moritz.riebe@mz-entwickler.de)
 * @param   <TPropertyValueType>    Typ des Wertes der Einstellung
 */
public interface OfficePropertyType<TPropertyValueType> {

    /**
     * Rückgabe des Namens dieser Einstellung.
     * 
     * <p>Der zurückgebene Name sollte identisch sein mit dem Namen der
     * statischen Variableninstanz.</p>
     * 
     * @return      z.B. {@code 'ERR_ON_VER_MISMATCH'}
     */
    public String name();
    
    /**
     * Überprüft den übergebenen Wert ob dieser valide ist für die Einstellung.
     * 
     * @param value     Wert der Einstellung, die gesetzt werden soll
     * 
     * @return          {@code true}, soweit dies eine zulässige
     *                  Einstellunsangabe ist.
     */
    public default boolean isValidPropertyValue(TPropertyValueType value) {
        return true;
    }
    
}
