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
package com.mz.solutions.office.model.interceptor;

import com.mz.solutions.office.OfficeDocumentException;

public abstract class ValueInterceptorException extends OfficeDocumentException {

    private ValueInterceptorException(String message) {
        super(message);
    }

    private ValueInterceptorException(String message, Throwable cause) {
        super(message, cause);
    }

    private ValueInterceptorException(Throwable cause) {
        super(cause);
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    public static class NoLowLevelSupportException extends ValueInterceptorException {
        
        NoLowLevelSupportException(String message) {
            super(message);
        }

        NoLowLevelSupportException(String message, Throwable cause) {
            super(message, cause);
        }

        NoLowLevelSupportException(Throwable cause) {
            super(cause);
        }
        
    }
    
    public static class NoXmlBasedDocumentException extends ValueInterceptorException {
        
        NoXmlBasedDocumentException(String message) {
            super(message);
        }

        NoXmlBasedDocumentException(String message, Throwable cause) {
            super(message, cause);
        }

        NoXmlBasedDocumentException(Throwable cause) {
            super(cause);
        }
        
    }
    
    public static class FailedInterceptorExecutionException extends ValueInterceptorException {
        
        public FailedInterceptorExecutionException(String message) {
            super(message);
        }

        public FailedInterceptorExecutionException(String message, Throwable cause) {
            super(message, cause);
        }

        public FailedInterceptorExecutionException(Throwable cause) {
            super(cause);
        }
        
    }

}
