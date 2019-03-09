/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2019,   Moritz Riebe     (moritz.riebe@mz-solutions.de)
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

/**
 * Package für die Verwendung von Value-Interceptoren.
 * 
 * <p>Jede Office-Implementierung unterstützt (derzeit) Value-Interceptors im Daten-Modell. Jeder
 * {@link ValueInterceptor} entspricht einem erweiterten Wert {@code ExtendedValue} und kann direkt
 * im Daten-Modell (z.B. {@code DataPage}, {@code DataTableRow}, ..) verwendet werden. Werte von
 * Interceptoren werden erst <b>bei Bedarf und Notwendigkeit</b> erzeugt und <b>nicht gecached</b>.
 * Value-Interceptors sind geeignet für aufwendig zu ladende, bauende Platzhalterwerte die wirklich
 * nur bei Notwendigkeit erzeugt werden sollen und/ oder Werte die in Abhängigkeit zum gewählten
 * Dokument/ Datensatz/ etc.. erzeugt werden müssen.</p>
 * 
 * <p>Für den Einstieg sollte die Dokumentation in {@link ValueInterceptor} und
 * {@link ValueInterceptorFunction} gelesen werden.</p>
 */
package com.mz.solutions.office.model.interceptor;
