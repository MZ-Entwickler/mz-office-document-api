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

/**
 * Einfache Platzhalter ohne Inhalt, aber mit indirekter Anweisung zur Layout-Formatierung des
 * Dokuments im Ersetzungsprozess um Tabellen und/oder Absätze dynamisch zu entfernen.
 * 
 * <p>In Vorlagen werden häufig je nach Fall einzelne Textbausteine eingefügt, oder entfernt. Ebenso
 * sind in manchen Fällen einzelne Tabellen nicht notwendig.</p>
 * 
 * <p>Formatierungs-Platzhalter sind <u>leere</u> Platzhalter (ohne Textinhalt) welche einer
 * Steuerungsfunktion zukommt. Solche Platzhalter können an den jeweiligen Stellen dazu führen das
 * Absätze entfernt werden, ganze Absätze "versteckt" werden oder Tabellen aus Dokumenten entfernt
 * werden welche im Ergebnis-Dokument nicht benötigt werden.</p>
 * 
 * <p>Ist einer der beiden Implementierungen jeweils einer dieser Steuerungsplatzhalter unbekannt,
 * wird ein leerer Platzhalter vermutet und die Dokumentenersetzung bricht nicht ab.</p>
 * 
 * Gegeben sei gedanklich zum Beispiel folgendes MS Word Dokument:
 * 
 * <pre>
 *  [Paragraph 1] Please pay the course registration fee now. The course fee is due two weeks before
 *                the course starts. { MERGEFIELD INV_P_NORMAL }
 *  [Paragraph 2] Registration fee and course fee are due two weeks before the course starts.
 *                { MERGEFIELD INV_P_TWO_WEEKS_ONLY }
 *  [Paragraph 3] Please pay the registration fee and course fee as quick as possible.
 *                { MERGEFIELD INV_P_LESS_THAN_TWO_WEEKS }
 * </pre>
 * 
 * Die einzelnen Absätze können durch diese Formatierungs-Platzhalter teilweise entfernt werden.
 * 
 * <pre>
 *  final DataPage invPage = new DataPage();
 *  final LocalDate courseBeginDate = ...
 *  final LocalDate invDate = ...
 * 
 *  final long dueDays = DAYS.between(invDate, courseBeginDate);
 * 
 *  if (dueDays &gt; 21) {
 *      invPage.addValue(new DataValue("INV_P_NORMAL", StandardFormatHint.PARAGRAPH_KEEP));
 *      invPage.addValue(new DataValue("INV_P_TWO_WEEKS_ONLY", StandardFormatHint.PARAGRAPH_REMOVE));
 *      invPage.addValue(new DataValue("INV_P_LESS_THAN_TWO_WEEKS", StandardFormatHint.PARAGRAPH_REMOVE));
 *  } else if (dueDays &gt; 14) {
 *      invPage.addValue(new DataValue("INV_P_NORMAL", StandardFormatHint.PARAGRAPH_REMOVE));
 *      invPage.addValue(new DataValue("INV_P_TWO_WEEKS_ONLY", StandardFormatHint.PARAGRAPH_KEEP));
 *      invPage.addValue(new DataValue("INV_P_LESS_THAN_TWO_WEEKS", StandardFormatHint.PARAGRAPH_REMOVE));
 *  } else {
 *      invPage.addValue(new DataValue("INV_P_NORMAL", StandardFormatHint.PARAGRAPH_REMOVE));
 *      invPage.addValue(new DataValue("INV_P_TWO_WEEKS_ONLY", StandardFormatHint.PARAGRAPH_REMOVE));
 *      invPage.addValue(new DataValue("INV_P_LESS_THAN_TWO_WEEKS", StandardFormatHint.PARAGRAPH_KEEP));
 *  }
 * </pre>
 */
package com.mz.solutions.office.model.hints;
