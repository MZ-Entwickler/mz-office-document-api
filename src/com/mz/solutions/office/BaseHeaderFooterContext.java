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
package com.mz.solutions.office;

import com.mz.solutions.office.instruction.HeaderFooterContext;

final class BaseHeaderFooterContext extends BaseInstructionContext implements HeaderFooterContext {

    private volatile boolean header, footer;
    private volatile String name;

    public BaseHeaderFooterContext(OfficeDocument document) {
        super(document);
    }
    
    public void setupAsHeader(String headerName) {
        this.header = true;
        this.footer = false;
        this.name = headerName;
    }
    
    public void setupAsFooter(String footerName) {
        this.header = false;
        this.footer = true;
        this.name = footerName;
    }
    
    @Override
    public boolean isHeader() {
        return header;
    }

    @Override
    public boolean isFooter() {
        return footer;
    }

    @Override
    public String getName() {
        return name;
    }
    
}
