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

/**
 * Office-spezifische Erweiterungen.
 * 
 * <p>Erweiterungen die entweder nur für ein Office implementiert wurden oder
 * nur in einem Office verfügbar sind können einzeln ergänzt werden. Die
 * entsprechenden Klassen/Schnittstellen finden sich hier.</p>
 * 
 * <p><b>Abrufen der Erweiterung (Schnittstelle):</b> Jede Erweiterung
 * implementiert als Marker die Schnittstelle {@link Extension}. Am eigentlichen
 * Office-Dokument kann abgefragt werden ob die Erweiterung unterstützt wird;
 * mit Rückgabe der Schnittstellen-Instanz.</p>
 * 
 * <pre>
 *  Optional&lt;MicrosoftCustomXml&gt; opExtension = anyDocument
 *          .extension(MicrosoftCustomXml.class);
 * 
 *  opExtension.isPresent() == true     // Erweiterung verfügbar
 *  opExtension.isPresent() == false    // Erweiterung nicht verfügbar
 * </pre>
 * 
 * <p>Erweiterungen sind nicht allgemein für die Office-Implementierung
 * verfügbar, sondern können vom jeweiligen Dateityp jeweils unterschiedlich
 * nutzbar sein.</p>
 */
package com.mz.solutions.office.extension;
