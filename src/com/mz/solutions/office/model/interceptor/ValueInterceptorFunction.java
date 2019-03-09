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
package com.mz.solutions.office.model.interceptor;

import com.mz.solutions.office.extension.ExtendedValue;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.model.DataValueMap;

/**
 * Funktionale Schnittstelle für die Methoden-Signatur eines Value-Interceptors.
 * 
 * @author Riebe, Moritz      (moritz.riebe@mz-solutions.de)
 */
@FunctionalInterface
public interface ValueInterceptorFunction {
   
    /**
     * Evaluierungs-Methode für den Platzhalterwert.
     * 
     * <p>Diese Methode erzeugt den Platzhalterwert sobald dieser in einem Dokument gefunden und
     * ersetzt werden muss. Die Rückgabe wird nicht automatisch gecached; eventuelles Caching
     * (z.B. von aufwendigeren Datenbankaktionen) muss bei Bedarf selbst übernommen werden.</p>
     * 
     * <p>Der Ersetzungsprozess kann durch Werfen von {@link ValueInterceptorException}, genauer
     * dessen Unterklasse {@link ValueInterceptorException.FailedInterceptorExecutionException},
     * als fehlgeschlagen angezeigt werden; entsprechend endet der gesamte Ersetzungsvorgang für
     * das Dokument.</p>
     * 
     * <p>Alle notwendigen Informationen (Name des Platzhalters, Dokument, Office-Implementierung
     * (Factory), andere Platzhalter im selben Scope ({@link DataValueMap})) für die Erzeugung des
     * ersetzungswertes können aus dem übergebenen {@code context} bezogen werden. Die Kontext-
     * Instanz darf nicht zwischengespeichert werden.</p>
     * 
     * @param context   Kontext für die Erzeugung des Platzhalterwertes. Ist nie {@code null}.
     * 
     * @return          Ergebnis das an die Office-Implementierung übergeben wird um den
     *                  Ersetzungsvorgang fortzuführen. Der Wert in {@link DataValueResult}, also
     *                  wieder ein {@link DataValue}, kann wiederum verschachtelt werden und selbst
     *                  ein erweiterter Wert ({@link ExtendedValue}) sein, ebenso ein erneuter
     *                  {@link ValueInterceptor}. Es darf <b>nicht</b> {@code null} zurück gegeben
     *                  werden! Die Name des Platzhalters im {@link DataValue} sollte möglichst
     *                  dem im Kontext übergebenen Platzhalternamen entsprechen; dies ist jedoch
     *                  nicht zwingend notwendig. Das Verhalten eines im Ergebnis abweichenden
     *                  Platzhalternamens ist aber unbestimmt.
     * 
     * @throws          ValueInterceptorException 
     *                  Signalisiert einen Fehler bei er Erzeugung des Ersetzungswertes und bringt
     *                  den gesamten Ersetzungsprozess zum Abbruch.
     */
    public DataValueResult interceptPlaceholder(InterceptionContext context)
            throws ValueInterceptorException;
    
}
