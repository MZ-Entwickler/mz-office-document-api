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
package com.mz.solutions.office.extension;

/**
 * Deklariert eine Klasse/ Komponente als office-spezifische Erweiterung die von
 * anderen Formaten/ Implementierungen nicht umgesetzt wurde/ umgesetzt werden
 * kann.
 * 
 * <p>Die Implementierte Erweiterung muss zwar diese Schnittstelle als Marker
 * implementieren, jedoch nicht umbedingt als eigene Klasse umgesetzt sein.</p>
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-entwickler.de)
 */
public interface Extension { }
