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

import com.mz.solutions.office.model.DataMap;
import java.util.Optional;

/**
 * Schnittstelle zur Entscheidung des Ersetzungs-Vorganges bei Kopf- und Fußzeilen im Dokument.
 * 
 * <p>Die Methode {@link #processHeaderFooter(HeaderFooterContext)} wird für jede im Dokument
 * gefundenen Kopf- und Fußzeile aufgerufen und befragt <u>ob</u> diese ersetzt werden soll und
 * wenn ja, mit welchen Daten der Ersetzungsvorgang ausgeführt werden soll.</p>
 * 
 * <p>Der übergeben {@link HeaderFooterContext} liefert zur Entscheidungsfindung weitere
 * Informationen wie Bezeichnung und ob es sich um eine Kopf- oder eine Fußzeile handelt.</p>
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-entwickler.de)
 */
@FunctionalInterface
public interface HeaderFooterInstruction extends DocumentProcessingInstruction {
    
    /**
     * Rückfrage beim Ersetzungs-Vorgang ob die gefundene Kopf- oder Fußzeile bei der Ersetzung
     * mit berücksichtigt wurde und mit welchen Daten die Ersetzung erfolgen soll.
     * 
     * <p>Diese Methode wird für jede Kopf- und Fußzeile <u>mindestens</u> einmal aufgerufen.
     * Besitzt eine Kopf- oder Fußzeile mehr als eine Bezeichnung (z.B. interne Bezeichnung,
     * Anzeige-Bezeichnung) wird diese Methode für jede Bezeichnung aufgerufen. Wurde bereits bei
     * der ersten Bezeichnung der Ersetzung zugestimmt, erfolgt kein folgender Aufruf für die
     * selbe Zeile mit zweiter Bezeichnung.</p>
     * 
     * @param context   Kontext zur Kopf- und Fußzeile mit Information der Zeilenart (Kopf|Fuß)
     *                  und Angabe des/eines Bezeichners.
     * 
     * @return          Rückgabe der Daten-Menge für den Ersetzungs-Vorgang, wenn dieser gewünscht
     *                  wird, oder Rückgabe von {@code Optional.empty()} wenn für die nachgefragte
     *                  Zeile keine Ersetzung stattfinden soll.
     */
    public Optional<DataMap<?>> processHeaderFooter(HeaderFooterContext context);
    
}
