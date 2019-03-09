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
package com.mz.solutions.office.extension;

import com.mz.solutions.office.model.DataValue;
import java.io.Serializable;

/**
 * Deklariert einen Datentypen der nicht eindeutig einer Zeichenkette zuordbar ist als erweiterten
 * Wert für einen Platzhalter; zB ein Bild, ein anderes Dokument oder andere Formen and Inhalten die
 * von Erweiterungen bereitgestellt werden.
 * 
 * <p>Klassen die hier von erben, können bei {@link DataValue} übergeben werden,
 * um einen alternativen Inhalt anzubieten.</p>
 * 
 * <p><b>Wichtiger Implementierungshinweis:</b> Unabhängig ob die Erweiterung genutzt werden kann
 * oder nicht, muss der Erweiterungs-Wert unbedingt eine Zeichenketten-Repräsentation als
 * Alternative unterstützen! Beide Methoden ({@link #altString()} und {@link #toString()}) sollten
 * keine performance-kritischen Aktionen ausführen und möglichst die selbe Rückgabe tätigen.!</p>
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
public abstract class ExtendedValue implements Serializable {
    
    /**
     * Gibt den alternativen Wert als Zeichenkette zurück (zB der Pfad zum Bild) um die
     * Kompatibilität zur anderen Dokumenten und Office Anwendungen zu erhalten.
     * 
     * @return      Zeichenkette, darf leer sein aber <b>nie</b> {@code null}.
     */
    public abstract String altString();
    
    /**
     * Repräsentation dieses erweiterten Wertes als (alternative!) Zeichenkette.
     * 
     * @return      Entspricht der Rückgabe von {@link #altString()}.
     */
    @Override
    public String toString() {
        return altString();
    }
    
}
