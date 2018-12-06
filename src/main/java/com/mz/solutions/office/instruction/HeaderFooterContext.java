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
package com.mz.solutions.office.instruction;

/**
 * Kontext beim Ersetzungs-Vorgang für die Entscheidungsfindung zur Ersetzung von einer oder
 * mehreren Kopf- und Fußzeilen.
 *
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
public interface HeaderFooterContext extends InstructionContext {

    /**
     * Prüfung ob es sich bei dieser Zeile um eine Kopfzeile handelt.
     *
     * @return {@code true}, wenn Kopfzeile
     */
    public boolean isHeader();

    /**
     * Prüfung ob es sich bei dieser Zeile um eine Fußzeile handelt.
     *
     * @return {@code true}, wenn Fußzeile
     */
    public boolean isFooter();

    /**
     * Rückgabe eines Namens oder Bezeichners der Kopf- oder Fußzeile.
     *
     * @return Bezeichner
     */
    public String getName();

}
