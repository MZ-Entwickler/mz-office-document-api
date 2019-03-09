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
package com.mz.solutions.office.instruction;

import com.mz.solutions.office.model.DataValueMap;
import java.util.List;

/**
 * Interceptor der vor und nach der Dokumentenbefüllung, einzelner Abschnittsweise, aufgerufen
 * werden kann.
 * 
 * <p>Wird bei der Ersetzung eine Instanz als Anweisung mit übergeben, wird jener der XML Baum
 * vor dem Ersetzungs-/Überarbeitungsprozess übergeben oder nach dem Prozess. Im XML gemachte
 * Änderungen werden dann direkt im Dokument angwendet und mit gespeichert. Bei Fehlern ist das
 * Ergebnis-Dokument ggf. nicht mehr anzeigbar.</p>
 * 
 * <p>Diese Schnittstelle kann händisch implementiert werden oder es kann die Vorimplementierung
 * aus der Klasse {@link DocumentInterceptor} (statische Fabriken) verwendet werden.</p>
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
public interface DocumentInterceptor extends DocumentProcessingInstruction, DocumentInterceptorFunction {
    
    /** Ersatz-Name um der Implementierung zur sagen das der Dokumenten-Inhalt gemeint ist. */
    public static final String GENERIC_PART_BODY = "#body";
    
    /** Ersatz-Name um der Implementierung zur sagen das der Formatierungen gemeint ist. */
    public static final String GENERIC_PART_STYLES = "#styles";
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    /**
     * Rückgabe des gewünschten Dokumenten-Abschnittes der abgefangen werden soll.
     * 
     * @return  Name des Abschnitts; Pfad im Dokument ohne führendes Slash-Zeichen,
     *          z.B. {@code '#body', 'META-INF/manifest.xml'}.
     */
    public String getPartName();
    
    /**
     * Rückgabe des Interceptor-Typus ob <u>vor</u> oder <u>nach</u> dem Ersetzungs-Vorgang dieser
     * Dokumenten-Interceptor aufgerufen werden soll.
     * 
     * @return  Instanz von {@link DocumentInterceptorType}.
     */
    public DocumentInterceptorType getInterceptorType();

    /**
     * Gibt die Daten zurück die speziell für die Ausführung des Interceptors gedacht sind.
     * 
     * @return  Daten für die Ausführung des Interceptors
     */
    public List<DataValueMap<?>> getInterceptorValues();
    
}
