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
package com.mz.solutions.office.model;

import com.mz.solutions.office.model.DataModelException.DataTableNameAlreadyExistsException;
import java.util.Optional;
import java.util.Set;

/**
 * Schnittstelle für Daten-Modelle die Untermengen als Tabellen enthalten
 * und diese zum <u>Abrufen</u> bereitstellen.
 * 
 * <p>Das Bearbeiten (Hinzufügen und Bennen von Tabellen) ist jedoch nur über
 * {@link DataTable} möglich.</p>
 * 
 * @author  Riebe, Moritz       (moritz.riebe@mz-entwickler.de)
 * 
 * @param   <TReturnInstance>   Typ für die Instanzrückgabe
 */
public interface DataTableMap<TReturnInstance> {
    
    /**
     * Fügt dem Daten-Modell eine benannte Tabelle hinzu.
     * 
     * @param table Tabellen-Daten-Modell; darf nicht {@code null} sein.
     * 
     * @return      Gibt die eigene Instanz zurück
     * 
     * @throws      DataTableNameAlreadyExistsException 
     *              Wenn der Name der Tabelle in der selben Ebene des
     *              Daten-Modells bereits vergeben ist.
     */
    public TReturnInstance addTable(DataTable table)
            throws DataTableNameAlreadyExistsException;
    
    /**
     * Sucht nach einer Tabelle mit dem übergebenen Tabellennamen.
     * 
     * @param tableName Name der Tabelle unabhängig der Groß- und
     *                  Kleinschreibung; darf nicht {@code null} sein.
     * 
     * @return          Tabelle mit der entsprechenden Bezeichnung
     *                  oder {@code Optional.empty()}
     */
    public Optional<DataTable> getTableByName(CharSequence tableName);
    
    
    /**
     * Gibt alle bisher hinzugefügten Tabellen zurück.
     * 
     * @return  Tabellen order {@code Collections.EMPTY_SET}; gibt niemals
     *          {@code null} zurück. Die zurückgegebene Menge ist nicht
     *          veränderbar.
     */
    public Set<DataTable> getTables();
    
}
