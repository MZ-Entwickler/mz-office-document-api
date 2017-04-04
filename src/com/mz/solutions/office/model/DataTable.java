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

import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import static mz.solutions.office.resources.DataTableKeys.DUP_TABLE_ROW;
import static mz.solutions.office.resources.DataTableKeys.TAB_NAME_TOO_SHORT;
import static mz.solutions.office.resources.MessageResources.formatMessage;

/**
 * Benanntes Tabelle mit einzelnen Bezeichner-Werte-Paaren und iterierbaren
 * Datensätzen.
 * 
 * <p>Kann zwei Formen von Einträgen besitzen. Einfache Platzhalter die
 * unabhängig von Datenzeilen ersetzt werden und sortierte Datenzeilen
 * die den Tabelleninhalt enthalten.</p>
 * 
 * @author  Riebe, Moritz       (moritz.riebe@mz-solutions.de)
 */
public final class DataTable extends AbstractDataSet
        implements Iterable<DataTableRow>, DataValueMap<DataTable> {
    
    private final String tableName;
    private final List<DataTableRow> tableRows = new LinkedList<>();
    
    /**
     * Erzeugt eine neue Tabelle mit dem übergebenen Tabellennamen.
     * 
     * @param tableName     Name der Tabelle; unabhängig von Groß- und
     *                      Kleinschreibung; muss aus mindestens 2 Zeichen
     *                      bestehen.
     */
    public DataTable(String tableName) {
        Objects.requireNonNull(tableName, "tableName");
        
        final String pTableName = tableName.trim().toUpperCase();
        
        if (pTableName.isEmpty() || pTableName.length() < 2) {
            throw new IllegalArgumentException(formatMessage(TAB_NAME_TOO_SHORT));
        }
        
        this.tableName = pTableName;
    }
    
    /**
     * Gibt den Tabellennamen in Großbuchstaben zurück.
     * 
     * @return      Name der Tabelle; nie {@code null}
     */
    public String getTableName() {
        return tableName;
    }
    
    /**
     * Fügt dem Model eine weitere Datenzeile hinzu.
     * 
     * @param tableRow  Anzufügende Zeile; darf nicht {@code null} sein;
     *                  selbe Instanzen einer Datenzeile dürfen nicht
     *                  mehrfach hinzugefügt werden!
     * 
     * @return          Gibt die eigene Instanz zurück
     * 
     * @throws          IllegalArgumentException
     *                  Wenn die übergebene Zeileninstanz bereits vorhanden ist
     */
    public DataTable addTableRow(DataTableRow tableRow) {
        Objects.requireNonNull(tableRow, "tableRow");
        
        if (tableRows.contains(tableRow)) {
            throw new IllegalArgumentException(formatMessage(DUP_TABLE_ROW,
                    /* {0} */ getClass().getSimpleName()));
        }
        
        tableRows.add(tableRow);
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Iterator<DataTableRow> iterator() {
        return tableRows.iterator();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public DataTable addValues(DataValue... values) {
        super.addValues(values);
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public DataTable addValue(DataValue value) {
        super.addValue(value);
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Optional<DataValue> getValueByKey(String keyName) {
        return super.getValueByKey(keyName);
    }
    
    /**
     * {@inheritDoc}
     */
    @Override
    public Set<DataValue> getValues() {
        return super.getValues();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Map<String, String> toMap() {
        return super.toMap();
    }
    
}
