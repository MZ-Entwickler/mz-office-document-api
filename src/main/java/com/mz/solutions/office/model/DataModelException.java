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

import com.mz.solutions.office.resources.DataModelExceptionKeys;
import com.mz.solutions.office.resources.MessageResources;

/**
 * Oberklasse aller Exceptions die im Umgang mit dem Daten-Modell auftreten.
 *
 * @author Riebe, Moritz       (moritz.riebe@mz-solutions.de)
 */
public abstract class DataModelException extends RuntimeException {

    protected DataModelException() {
    }

    protected DataModelException(String message) {
        super(message);
    }

    protected DataModelException(String message, Throwable cause) {
        super(message);
        addSuppressed(cause);
    }

    protected DataModelException(Throwable cause) {
        this(cause.getMessage(), cause);
    }

    ////////////////////////////////////////////////////////////////////////////

    /**
     * Tritt auf wenn Platzhalterbezeichner bereits vergeben sind um
     * Duplikate zu vermeiden.
     */
    public static class DataValueKeyNameAlreadyExistsException
            extends DataModelException {

        protected DataValueKeyNameAlreadyExistsException(
                final Class<?> modelClass, final String keyName) {

            super(MessageResources.formatMessage(DataModelExceptionKeys.DUP_KEY_NAME,
                    /* {0} */ modelClass.getSimpleName(),
                    /* {1} */ keyName,
                    /* {2} */ DataValue.class.getSimpleName()));
        }

    }

    /**
     * Tritt auf wenn ein Tabellenname bereits vergeben ist.
     */
    public static class DataTableNameAlreadyExistsException
            extends DataModelException {

        protected DataTableNameAlreadyExistsException(
                final Class<?> modelClass, final String tableName) {

            super(MessageResources.formatMessage(DataModelExceptionKeys.DUP_TABLE_NAME,
                    /* {0} */ modelClass.getSimpleName(),
                    /* {1} */ tableName));
        }

    }
}
