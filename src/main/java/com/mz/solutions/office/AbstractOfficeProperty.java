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

import java.util.Objects;

/**
 * Intern - Oberklasse für die Definition von verfügbaren OfficePropertie's.
 *
 * <p>Da sich die Implementierung mehrerer Methoden für die Microsoft und
 * OpenOffice Implementierung der Einstellungen wiederholen, wurde dies
 * die gemeinsame Oberklasse.</p>
 *
 * @param <TPropertyValueType> Typ des Einstellungsparameters
 * @author Riebe, Moritz   (moritz.riebe@mz-solutions.de)
 */
abstract class AbstractOfficeProperty<TPropertyValueType>
        implements OfficePropertyType<TPropertyValueType> {

    private final String name;

    protected AbstractOfficeProperty(String name) {
        this.name = name;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String name() {
        return name;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isValidPropertyValue(TPropertyValueType value) {
        return null != value;
    }

    @Override
    public int hashCode() {
        return Objects.hash(AbstractOfficeProperty.class, this.name);
    }

    @Override
    public boolean equals(Object obj) {
        if (obj == null || obj instanceof OfficePropertyType == false) {
            return false;
        }

        if (getClass() != obj.getClass()) {
            return false;
        }

        final AbstractOfficeProperty<?> other = (AbstractOfficeProperty<?>) obj;
        return Objects.equals(this.name, other.name);
    }

    /**
     * Gibt den Klassennamen zurück sowie den Namen der Einstellung.
     *
     * @return String-Repräsentation der Einstellung; ohne Wert
     */
    @Override
    public String toString() {
        return getClass().getSimpleName() + "." + name;
    }

}
