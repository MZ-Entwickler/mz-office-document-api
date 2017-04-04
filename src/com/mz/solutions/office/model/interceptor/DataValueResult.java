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
package com.mz.solutions.office.model.interceptor;

import com.mz.solutions.office.model.DataValue;
import java.io.Serializable;
import java.util.Objects;
import javax.annotation.CheckForNull;
import javax.annotation.Nullable;

/**
 * Ergebnis eines {@link ValueInterceptor}'s nach dessen Ersetzungswert erzeugung.
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
public final class DataValueResult implements Serializable {
    
    /**
     * Für die Ersetzung soll der im Parameter übergebene Wert verwendet werden.
     * 
     * @param value     Wert für die Ersetzung, darf nicht {@code null} sein.
     * 
     * @return          Instanz eines Verarbeitungsergebnisses nach dem Durchlaufen eines
     *                  {@link ValueInterceptor}'s.
     */
    public static final DataValueResult useValue(DataValue value) {
        return new DataValueResult(value);
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////

    private final DataValue value;
    
    private DataValueResult(DataValue value) {
        this.value = Objects.requireNonNull(value, "(DataValue) value");
    }
    
    @Nullable @CheckForNull
    public DataValue value() {
        return value;
    }
    
}
