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
package com.mz.solutions.office.instruction;

/**
 * Art des Document-Interceptors.
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-entwickler.de)
 */
public enum DocumentInterceptorType {
    
    /** Aufrufen vor dem Ersetzungs-Prozess. */
    BEFORE_GENERATION,
    
    /** Aufrufen nach dem Ersetzungs-Prozess. */
    AFTER_GENERATION
    
}
