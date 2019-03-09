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
package com.mz.solutions.office;

import com.mz.solutions.office.extension.ExtendedValue;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.model.interceptor.DataValueResult;
import com.mz.solutions.office.model.interceptor.InterceptionContext;
import com.mz.solutions.office.model.interceptor.ValueInterceptor;
import com.mz.solutions.office.model.interceptor.ValueInterceptorException;
import javax.annotation.CheckForNull;
import javax.annotation.Nullable;

final class ValueInterceptorUtil {
    
    @Nullable @CheckForNull
    public static final DataValue callInterceptors(
            final DataValue dataValue,
            final InterceptionContext interceptionContext)
            throws ValueInterceptorException
    {
        if (dataValue.isExtendedValue() == false) {
            return dataValue;
        }
        
        final ExtendedValue extendedValue = dataValue.extendedValue();
        
        if ((extendedValue instanceof ValueInterceptor) == false) {
            return dataValue;
        }
        
        final ValueInterceptor valueInterceptor = (ValueInterceptor) extendedValue;
        final DataValueResult interceptorResult;
        
        try {
            interceptorResult = valueInterceptor.interceptPlaceholder(interceptionContext);
            
        } catch (ValueInterceptorException valueInterceptorException) {
            throw valueInterceptorException;
        } catch (Exception otherException) {
            throw new ValueInterceptorException.FailedInterceptorExecutionException(otherException);
        }
        
        final DataValue resultValue = interceptorResult.value();
        
        if (null == resultValue) {
            return null;
        }
        
        return callInterceptors(resultValue, interceptionContext);
    }
    
}
