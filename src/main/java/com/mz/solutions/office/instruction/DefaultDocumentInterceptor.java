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

import com.mz.solutions.office.model.DataValueMap;

import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/**
 * Interne Standard-Implementierung für die Factory-Methoden.
 * 
 * @author Riebe, Moritz      (moritz.riebe@mz-entwickler.de)
 */
final class DefaultDocumentInterceptor implements DocumentInterceptor {

    private final String partName;
    private final DocumentInterceptorType interceptorType;
    private final DocumentInterceptorFunction function;
    private final DataValueMap[] moreValues;

    public DefaultDocumentInterceptor(
            String partName, DocumentInterceptorType interceptorType,
            DocumentInterceptorFunction function, DataValueMap ... moreValues)
    {
        this.partName = Objects.requireNonNull(partName, "partName");
        this.interceptorType = Objects.requireNonNull(interceptorType, "interceptorType");
        this.function = Objects.requireNonNull(function, "function");
        this.moreValues = Objects.requireNonNull(moreValues, "moreValues");
    }
    
    @Override
    public String getPartName() {
        return partName;
    }

    @Override
    public DocumentInterceptorType getInterceptorType() {
        return interceptorType;
    }

    @Override
    public List<DataValueMap<?>> getInterceptorValues() {
        return Arrays.asList(moreValues);
    }

    @Override
    public void callInterceptor(DocumentInterceptionContext context) {
        function.callInterceptor(context);
    }
    
}
