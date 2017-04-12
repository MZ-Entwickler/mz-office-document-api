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

import java.io.Serializable;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import javax.annotation.Nullable;

/**
 * Intern - Basisimplementierung zum Speichern und Abfragen von
 * Bezeichner-Werte-Paaren und Untertabellen.
 * 
 * @author  Riebe, Moritz   (moritz.riebe@mz-solutions.de)
 */
abstract class AbstractDataSet implements Serializable {
    
    // values == null := keine Werte hinterlegt/eingerichtet
    private List<DataValue> values;
    
    // tables == null := keine Tabellen hinterlegt/eingerichtet
    private LinkedList<DataTable> tables;
    
    private void initValueLists() {
        if (null == values) {
            values = new LinkedList<>();
        }
    }
    
    final String safeToString(@Nullable CharSequence charSeq, String varName) {
        if (null == charSeq) {
            throw new NullPointerException(varName);
        }
        
        final String strValue = charSeq.toString();
        if (null == strValue) {
            throw new NullPointerException(varName + ".toString()");
        }
        
        return strValue;
    }
    
    protected Optional<DataValue> getValueByKey(CharSequence keyName) {
        Objects.requireNonNull(keyName, "keyName");
        
        if (null == values) {
            return Optional.empty();
        }
        
        final String pKeyName = safeToString(keyName, "keyName").trim().toUpperCase();
        
        if (pKeyName.isEmpty() || pKeyName.length() < 2) {
            return Optional.empty();
        }
        
        final Iterator<DataValue> valueIterator = values.iterator();
        while (valueIterator.hasNext()) {
            final DataValue value = valueIterator.next();
            
            if (value.getKeyName().equals(pKeyName)) {
                return Optional.of(value);
            }
        }
        
        return Optional.empty();
    }
    
    protected Set<DataValue> getValues() {
        if (null == values) {
            return Collections.EMPTY_SET;
        }
        
        return Collections.unmodifiableSet(new HashSet<>(values));
    }
    
    protected Map<String, String> toMap() {
        if (null == values || values.isEmpty()) {
            return Collections.EMPTY_MAP;
        }
        
        final Map<String, String> resultMap = new HashMap<>(values.size());
        
        for (DataValue value : values) {
            resultMap.put(value.getKeyName(), value.getValue());
        }
        
        return Collections.unmodifiableMap(resultMap);
    }
    
    protected AbstractDataSet addValue(DataValue value) {
        Objects.requireNonNull(value, "value");
        
        final String keyName = value.getKeyName();
        final Optional<DataValue> existingValue = getValueByKey(keyName);
        
        if (existingValue.isPresent()) {
            throw new DataModelException
                    .DataValueKeyNameAlreadyExistsException(
                            getClass(), keyName);
        }
        
        initValueLists();
        values.add(value);
        
        return this;
    }
    
    protected AbstractDataSet addValues(DataValue ... values) {
        if (null == values || values.length == 0) {
            return this;
        }
        
        for (int i = 0; i < values.length; i++) {
            final DataValue value = values[i];
            
            if (null == value) {
                throw new NullPointerException("value[" + i + "] == null");
            }
            
            addValue(value);
        }
        
        return this;
    }
    
    private void initTableList() {
        if (null == tables) {
            tables = new LinkedList<>();
        }
    }
    
    protected AbstractDataSet addTable(DataTable table) {
        Objects.requireNonNull(table, "table");
        
        final String tableName = table.getTableName();
        
        if (null != tables && getTableByName(tableName).isPresent()) {
            throw new DataModelException
                    .DataTableNameAlreadyExistsException(getClass(), tableName);
        }
        
        initTableList();
        tables.add(table);
        
        return this;
    }
    
    protected Optional<DataTable> getTableByName(CharSequence tableName) {
        Objects.requireNonNull(tableName, "tableName");
        
        if (null == tables) {
            return Optional.empty();
        }
        
        final String pTableName = safeToString(tableName, "tableName").trim().toUpperCase();
        
        for (DataTable singleTable : tables) {
            if (singleTable.getTableName().equals(pTableName)) {
                return Optional.of(singleTable);
            }
        }
        
        return Optional.empty();
    }
    
    protected Set<DataTable> getTables() {
        if (null == tables) {
            return Collections.EMPTY_SET;
        }
        
        return Collections.unmodifiableSet(new HashSet<>(tables));
    }
    
}
