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
package com.mz.solutions.office.model.images;

/**
 * Längen-Einheiten und Umrechnungen für die Angabe von Bild-Abmaßungen (CM, MM, EMU).
 * 
 * <p>Es werden nur CM und MM direkt unterstützt.</p>
 * 
 * <pre>
 *  final double imageWidth = UnitOfLength.CENTIMETERS.toMillimeters(15.0D); // cm zu mm
 * </pre>
 * 
 * @see ImageValue
 *      Angabe von Bild-Abmaßungen.
 * 
 * @author Riebe, Moritz      (moritz.riebe@mz-entwickler.de)
 */
public enum UnitOfLength {
    
    MILLIMETERS () {
        /** {@inheritDoc} */ @Override
        public double toMillimeters(double length) { return length; }
        
        /** {@inheritDoc} */ @Override
        public double toCentimeters(double length) { return length / 10.0D; }
        
        /** {@inheritDoc} */ @Override
        public double toEnglishMetricUnits(double length) { return EMU_PER_MM * length; }
    },
    
    CENTIMETERS () {
        /** {@inheritDoc} */ @Override
        public double toMillimeters(double length) { return length * 10.0D; }
        
        /** {@inheritDoc} */ @Override
        public double toCentimeters(double length) { return length; }
        
        /** {@inheritDoc} */ @Override
        public double toEnglishMetricUnits(double length) { return EMU_PER_CM * length; }
    };
    
    private static final double EMU_PER_CM = 360000;
    private static final double EMU_PER_MM = EMU_PER_CM / 10.0D;
    
    private UnitOfLength() { }
    
    /**
     * Rechnet die übergebene Länger in Millimeter um.
     * 
     * @param length    Länge
     * 
     * @return          Länge im Millimeter.
     */
    public abstract double toMillimeters(double length);
    
    /**
     * Rechnet die übergebene Länge in Centimeter um.
     * 
     * @param length    Länge
     * 
     * @return          Länge in Centimeter.
     */
    public abstract double toCentimeters(double length);
    
    /**
     * Rechnet die übergebe Länge in Englisch-Metric-Units (emu) um.
     * 
     * @param length    Länge
     * 
     * @return          Länge in {@code EMU} für MS Word.
     */
    public abstract double toEnglishMetricUnits(double length);
    
}
