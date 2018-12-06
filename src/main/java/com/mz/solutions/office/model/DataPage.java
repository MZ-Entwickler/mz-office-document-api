/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2018,   Moritz Riebe     (moritz.riebe@mz-solutions.de),
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

import java.util.Map;
import java.util.Optional;
import java.util.Set;

/**
 * Entspricht der Sammlung aller Datensätze für ein Dokument/ Seite die
 * einzusetzen sind.
 *
 * @author Riebe, Moritz       (moritz.riebe@mz-solutions.de)
 */
public final class DataPage extends AbstractDataSet
        implements DataValueMap<DataPage>,
        DataTableMap<DataPage>,
        DataMap<DataPage> {

    /**
     * {@inheritDoc}
     */
    @Override
    public Optional<DataTable> getTableByName(CharSequence tableName) {
        return super.getTableByName(tableName);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public DataPage addTable(DataTable table) {
        super.addTable(table);
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Set<DataTable> getTables() {
        return super.getTables();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public DataPage addValues(DataValue... values) {
        super.addValues(values);
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public DataPage addValue(DataValue value) {
        super.addValue(value);
        return this;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Optional<DataValue> getValueByKey(CharSequence keyName) {
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
