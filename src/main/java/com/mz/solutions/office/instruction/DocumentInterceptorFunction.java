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
package com.mz.solutions.office.instruction;

/**
 * Callback-Schnittstelle für Dokumenten-Interceptoren.
 *
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
@FunctionalInterface
public interface DocumentInterceptorFunction {

    /**
     * Callback-Methode die aufgerufen wird, wenn der Zeitpunkt zum Bearbeiten des
     * Dokumenten-Abschnitts gekommen ist.
     *
     * @param context Kontext mit erweiterten Angaben.
     */
    public void callInterceptor(DocumentInterceptionContext context);

}
