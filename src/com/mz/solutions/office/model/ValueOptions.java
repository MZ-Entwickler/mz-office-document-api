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
package com.mz.solutions.office.model;

/**
 * Strategy zur Handhabung von Zeilenumbrüchen, Tabulatorzeichen und weiteren
 * Optionen in den Werten von Platzhaltern.
 * 
 * <p>Die Optionen können einzel gewählt werden bei der Erstellung eines jeden
 * {@link DataValue}. Standardmäßig sind die Optionen so gewählt, das
 * Zeilchenumbrüche und Tabulator-Zeichen in den Werten beibehalten werden und 
 * korrekt ins Dokument übernommen werden.</p>
 * 
 * @author  Riebe, Moritz       (moritz.riebe@mz-solutions.de)
 */
public enum ValueOptions {
    
    // Beim Erweitern um Optionen müssen diese auch unten in den
    // #isTabulatorOption() und #isLineBreakOption() vermerkt werden!
    
    /**
     * Umwandln von Zeilenumbrüchen in Leerzeichen.
     * 
     * <p>Konvertiert {@code '\n'} als Leerzeichen und entfernt alle
     * {@code '\r'} aus der Zeichenkette.</p>
     */
    LINEBREAK_TO_WHITESPACE,
    
    /**
     * Beibehalten von Zeilenumbrüchen und als Formatierungsbaum im
     * OpenDocument umsetzen.
     * 
     * <p>Entfernt zwar alle Zeilenumbrüche, setzt diese jedoch im Dokument
     * als Formatierungsangabe mit um. Zeilenumbrüche sind dabei innerhalb
     * eines Absatzes und erben die Formatierung des Eltern-Absatzes.</p>
     */
    KEEP_LINEBREAK,
    
    /**
     * Ignoriert enthaltene Tabulatoren und wandelt diese in ein Leerzeichen um.
     */
    TABULATOR_TO_WHITESPACE,
    
    /**
     * Konvertiert Tabulatorzeichen zu 4 Leerzeichen.
     */
    TABULATOR_TO_FOUR_SPACES,
    
    /**
     * Beibehalten von Tabulatoren und als Formatierung im Dokument umsetzen.
     */
    KEEP_TABULATOR;
    
    /**
     * Interne Funktion - Abfrage ob es sich um eine Tabulator-Option handelt.
     * 
     * @return  {@code true}, wenn es eine Tabulator-Option ist
     */
    protected boolean isTabulatorOption() {
        return this == TABULATOR_TO_FOUR_SPACES
                || this == TABULATOR_TO_WHITESPACE
                || this == KEEP_TABULATOR;
    }
    
    /**
     * Interne Funktion - Abfrage ob es eine Zeilenumbruchs-Option ist.
     * 
     * @return  {@code true}, wenn Option Zeilenumbrüche behandelt
     */
    protected boolean isLineBreakOption() {
        return this == KEEP_LINEBREAK
                || this == LINEBREAK_TO_WHITESPACE;
    }
    
}
