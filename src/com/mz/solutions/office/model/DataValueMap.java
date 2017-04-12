/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2017,   Moritz Riebe     (moritz.riebe@mz-solutions.de)
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

import com.mz.solutions.office.model.DataModelException.DataValueKeyNameAlreadyExistsException;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

/**
 * Schnittstelle für Daten-Modelle die Bezeichner-Wert-Beziehungen anbieten,
 * beinhalten und verwalten.
 * 
 * <p>Container die {@link DataValue}'s anbieten, implementieren diese
 * Schnittstelle.</p>
 * 
 * @see     DataMap
 *          Modelle die neben Bezeichner-Wert-Beziehungen auch Untermengen
 *          (Tabellen) implementieren.
 * 
 * @see     DataTableMap
 *          Modelle die Untermengen (Tabellen) implementieren.
 * 
 * @author  Riebe, Moritz       (moritz.riebe@mz-solutions.de)
 * 
 * @param <TReturnInstance>     Intern - Typ für Instanzrückgabe
 */
public interface DataValueMap<TReturnInstance> {
    
    /**
     * Sucht nach dem Wert eines Platzhalters anhand dessen Bezeichnung.
     * 
     * @param keyName   Name des Platzhalters; Groß- und Kleinschreibung sind
     *                  nicht relevant; darf nicht {@code null} sein; muss
     *                  mindestens 2 Zeichen lang sein, ohne Leerzeichen.
     * 
     * @return          Wert des Platzhalters als {@link DataValue} oder
     *                  {@code Optional.empty()} sollte kein Wert gefunden
     *                  worden sein.
     */
    public Optional<DataValue> getValueByKey(CharSequence keyName);
    
    /**
     * Gibt alle auf dieser Ebene eingetragenen Wertpaare zurück.
     * 
     * @return          Die zurückgegebene Menge ist nicht bearbeitbar.
     *                  Es wird nie {@code null} zurückgegeben.
     */
    public Set<DataValue> getValues();
    
    /**
     * Gibt alle Ersetzungspaare als normale (read-only) Map zurück.
     * 
     * @return          Rückgabe als Map; niemals {@code null} und die zurück
     *                  gegebene Map ist nicht bearbeitbar (read-only).
     */
    public Map<String, String> toMap();
    
    /**
     * Fügt diesem Modell eine neuen Platzhalter hinzu.
     * 
     * @param value Platzhalter, darf nicht {@code null} sein.
     * 
     * @return      Gibt die eigene Instanz zurück
     * 
     * @throws      DataValueKeyNameAlreadyExistsException
     *              Wenn der Platzhalterbezeichner bereits vergeben/ belegt ist.
     */
    public TReturnInstance addValue(DataValue value)
            throws DataValueKeyNameAlreadyExistsException;
    
    /**
     * Fügt diesem Modell einen oder mehrere Platzhalter hinzu.
     * 
     * <p>Verhält sich identisch zum mehrfachen Aufruf von
     * {@link #addValue(DataValue)}.</p>
     * 
     * @param values    Menge aller Platzhalter die diesem Modell hinzugefügt
     *                  werden sollen; muss mindestens 1 Platzhalter sein und
     *                  darf nicht {@code null} sein.
     * 
     * @return          Gibt die eigene Instanz zurück
     * 
     * @throws          DataValueKeyNameAlreadyExistsException 
     *                  Wenn einer der übergebenen Platzhalter einen Bezeichner
     *                  besitzt der bereits vergeben/ belegt ist.
     */
    public TReturnInstance addValues(DataValue ... values)
            throws DataValueKeyNameAlreadyExistsException;
    
}
