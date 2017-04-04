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

/**
 * Einstellungen der Factories die für alle Office Implementierungen existieren.
 * 
 * <p>Alle verfügbaren Einstellung sind als statische Attribute angelegt und
 * enthalten in deren Beschreibung die Voreinstellung sowie deren
 * Verhaltensbeschreibung.</p>
 * 
 * @author  Riebe, Moritz       (moritz.riebe@mz-solutions.de)
 * 
 * @param <TPropertyValue>      Typ des anzunehmenden Wertes
 */
public final class OfficeProperty<TPropertyValue>
        extends AbstractOfficeProperty<TPropertyValue> {
    
    /**
     * Verhaltenseinstellung wie mit Dokumenten die in einem "bekannten"
     * Format vorliegen, jedoch mit einer neueren Office Version angelegt
     * wurde, umgegangen werden soll.
     * 
     * <p>Die Voreinstellung ist {@code Boolean.FALSE}; neuere Dokumentenformate
     * werden soweit wie möglich behandelt.</p>
     */
    public static final OfficeProperty<Boolean> ERR_ON_VER_MISMATCH;
    
    /**
     * Fehler bei fehlendem Wert zu einem Platzhalter im Dokument.
     * 
     * <p>Existiert zu einem Platzhalter kein entsprechender Wert wird bei der
     * Voreinstellung ({@code Boolean.TRUE}) der Vorgang abgebrochen.</p>
     */
    public static final OfficeProperty<Boolean> ERR_ON_MISSING_VAL;
    
    /**
     * Verhalten wenn bei der Dokumentenerstellung kein Daten-Modell existiert.
     * 
     * <p>Bei {@code Boolean.TRUE} wird eine
     * {@link OfficeDocumentException.NoDataForDocumentGenerationException}
     * geworfen, sobald die Menge des Iterators 0 ist. <br><br>
     * 
     * Bei {@code Boolean.FALSE} wird keine Exception geworfen und auch kein
     * Dokument angelegt. Der Vorgang wird leise beendet. <br><br>
     * 
     * Die Voreinstellug ist {@code Boolean.TRUE}.</p>
     */
    public static final OfficeProperty<Boolean> ERR_ON_NO_DATA;
    
    static {
        ERR_ON_VER_MISMATCH = new OfficeProperty<>("ERR_ON_VER_MISMATCH");
        ERR_ON_MISSING_VAL = new OfficeProperty<>("ERR_ON_MISSING_VAL");
        ERR_ON_NO_DATA = new OfficeProperty<>("ERR_ON_NO_DATA");
    }
    
    ////////////////////////////////////////////////////////////////////////////
    
    private OfficeProperty(String name) {
        super(name);
    }

    @Override
    public boolean isValidPropertyValue(TPropertyValue value) {
        return value instanceof Boolean;
    }

}
