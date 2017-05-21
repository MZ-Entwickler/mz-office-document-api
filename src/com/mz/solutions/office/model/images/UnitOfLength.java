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
package com.mz.solutions.office.model.images;

public enum UnitOfLength {
    
    MILLIMETERS () {
        public double toMillimeters(double length) { return length; }
        public double toCentimeters(double length) { return length / 10.0D; }
        public double toEnglishMetricUnits(double length) { return EMU_PER_MM * length; }
    },
    
    CENTIMETERS () {
        public double toMillimeters(double length) { return length * 10.0D; }
        public double toCentimeters(double length) { return length; }
        public double toEnglishMetricUnits(double length) { return EMU_PER_CM * length; }
    };
    
    private static final double EMU_PER_CM = 360000;
    private static final double EMU_PER_MM = EMU_PER_CM / 10.0D;
    
    private UnitOfLength() { }
    
    public abstract double toMillimeters(double length);
    public abstract double toCentimeters(double length);
    
    public abstract double toEnglishMetricUnits(double length);
    
}
