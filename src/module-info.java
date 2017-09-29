/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2016,   Moritz Riebe     (moritz.riebe@mz-solutions.de)
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
module com.mz.solutions.office {
    exports com.mz.solutions.office;
    exports com.mz.solutions.office.extension;
    exports com.mz.solutions.office.instruction;
    exports com.mz.solutions.office.model;
    exports com.mz.solutions.office.model.images;
    exports com.mz.solutions.office.model.interceptor;
    exports com.mz.solutions.office.result;
    exports com.mz.solutions.office.util;
    
    requires java.logging;
    requires java.xml;
    
    requires jsr305.javax.annotation;
}
