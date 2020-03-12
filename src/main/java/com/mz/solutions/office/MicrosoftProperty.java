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
package com.mz.solutions.office;

import com.mz.solutions.office.model.DataPage;

/**
 * Spezifische Einstellungen für Microsoft Office Dokumente.
 * 
 * <p>Die Einstellungen sind nur auf die Implementierung für
 * Microsoft Dokumente anwendbar.</p>
 * 
 * @author      Riebe, Moritz       (moritz.riebe@mz-entwickler.de)
 * 
 * @param <TPropertyValue>      Typ den der Wert der Einstellung haben muss
 */
public final class MicrosoftProperty<TPropertyValue>
        extends AbstractOfficeProperty<TPropertyValue> {

    /**
     * Verhalten wenn mehrere Datensätze (Seiten) für ein Dokument vorliegen.
     * 
     * <p>Bei der Voreinstellung {@code Boolean.TRUE} wird zwischen jeder
     * {@link DataPage} explizit ein Seitenumbruch eingefügt. Bei
     * {@code Boolean.FALSE} werden keinerlei Seitenumbruche eingefügt; das
     * Verhalten ergiebt sich ausschließlich aus den Absatzeinstellungen.</p>
     */
    public static final MicrosoftProperty<Boolean> INS_HARD_PAGE_BREAKS;
    
    /**
     * Bestimmt das einzufügende Bilder - soweit möglich - per Drawing-Element ins Dokument
     * eingefügt werden sollen und nicht per VML (VML Shape).
     * 
     * <p>Bei der Voreinstellung {@code Boolean.TRUE} wird, soweit dies funktional möglich ist,
     * das Drawing-Element (Zeichnungs-Element) verwendet um Bilder/Grafiken in Word-Dokumenten
     * einzufügen. Wird auf {@code Boolean.FALSE} umgestellt, wird - soweit möglich - VML verwendet
     * um Bilder einzufügen.</p>
     */
    public static final MicrosoftProperty<Boolean> USE_DRAWING_OVER_VML;
    
    static {
        INS_HARD_PAGE_BREAKS = new MicrosoftProperty<>("INS_HARD_PAGE_BREAKS");
        USE_DRAWING_OVER_VML = new MicrosoftProperty<>("USE_DRAWING_OVER_VML");
    }
    
    ////////////////////////////////////////////////////////////////////////////
    
    private MicrosoftProperty(String name) {
        super(name);
    }

    @Override
    public boolean isValidPropertyValue(TPropertyValue value) {
        return value instanceof Boolean;
    }

}
