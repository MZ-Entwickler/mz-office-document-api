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
package mz.solutions.office.resources;

import java.text.MessageFormat;
import java.util.ResourceBundle;

/**
 * Interne Klasse zum Aufsuchen der richtigen Lokalisierung.
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
public final class MessageResources {
    private static final String PACKAGE_NAME = MessageResources.class.getPackage().getName();
    private static final String RES_LOCALE_FULL_PATH = PACKAGE_NAME + ".locale_office";
    
    private MessageResources() {
        throw new AssertionError("Kein " + getClass().getSimpleName() + " für dich!");
    }
    
    public static ResourceBundle getLocaleBundle() {
        return ResourceBundle.getBundle(RES_LOCALE_FULL_PATH);
    }
    
    public static String getString(String keyName) {
        return getLocaleBundle().getString(keyName);
    }
    
    public static String formatMessage(String keyName, Object ... args) {
        return MessageFormat.format(getString(keyName), args);
    }
    
}
