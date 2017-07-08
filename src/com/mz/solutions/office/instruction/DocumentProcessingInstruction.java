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
package com.mz.solutions.office.instruction;

import static com.mz.solutions.office.instruction.DocumentInterceptor.GENERIC_PART_BODY;
import static com.mz.solutions.office.instruction.DocumentInterceptor.GENERIC_PART_STYLES;
import static com.mz.solutions.office.instruction.DocumentInterceptorType.AFTER_GENERATION;
import static com.mz.solutions.office.instruction.DocumentInterceptorType.BEFORE_GENERATION;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataValueMap;
import java.util.Objects;
import java.util.Optional;
import javax.annotation.Nullable;

/**
 * Markierungs-Interface für erweiterte Anweisungen bei der Generierung/Erzeugung des
 * Ziel-Dokumentes.
 * 
 * <p>Einzelne Erweiterungen die nicht grundsätzlich notwendig sind, z.B. Ersetzung von Kopf- und
 * Fußzeilen, Office-spezifische Dinge etc... können als Menge von Dokumenten-Anweisungen bei der
 * Generierung mit übergeben werden.</p>
 * 
 * <p>Auf Dokument-Anwensungen aufbauend, werden auch Dokumenten-Listener implementiert.</p>
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
public interface DocumentProcessingInstruction {
    
    /**
     * Anweisung alle Kopf- und Fußzeilen im Ergebnis-Dokument mit den übergeben Werten zu ersetzen.
     * 
     * @param values    Ersetzungs-Werte
     * 
     * @return          Kopf-/Fußzeilen-Anweisung
     */
    public static HeaderFooterInstruction replaceHeaderFooterWith(DataPage values) {
        Objects.requireNonNull(values, "values");
        
        return headerFooterContext -> {
            return Optional.of(values);
        };
    }
    
    /**
     * Anweisung Kopf- und Fußzeilen im Ergebnis-Dokument mit den übergeben Werten zu ersetzen wenn
     * der Name der Kopf- und Fußzeile mit dem in {@code name} übergebenen übereinstimmt.
     * 
     * <p>Der Name der Kopf- oder Fußzeile ist unabhängig der Groß- und Kleinschreibung.</p>
     * 
     * @param name      Name Kopf-/Fußzeile
     * 
     * @param values    Ersetzungs-Werte
     * 
     * @return          Kopf-/Fußzeilen-Anweisung
     */
    public static HeaderFooterInstruction replaceHeaderFooterWith(String name, DataPage values) {
        Objects.requireNonNull(name, "name");
        Objects.requireNonNull(values, "values");
        
        return headerFooterContext -> {
            if (name.equalsIgnoreCase(headerFooterContext.getName())) {
                return Optional.of(values);
            } else {
                return Optional.empty();
            }
        };
    }
    
    /**
     * Anweisung alle Kopfzeilen im Ergebnis-Dokument zu ersetzen mit den übergebenen Werten.
     * 
     * @param values    Ersetzungs-Werte
     * 
     * @return          Kopfzeilen-Anweisung
     */
    public static HeaderFooterInstruction replaceHeaderWith(DataPage values) {
        Objects.requireNonNull(values, "values");
        
        return headerFooterContext -> {
            return headerFooterContext.isHeader() ? Optional.of(values) : Optional.empty();
        };
    }
    
    /**
     * Anweisung alle Fußzeilen im Ergebnis-Dokument zu ersetzen mit den übergeben Werten.
     * 
     * @param values    Ersetzungs-Werte
     * 
     * @return          Fußzeilen-Anweisung
     */
    public static HeaderFooterInstruction replaceFooterWith(DataPage values) {
        Objects.requireNonNull(values, "values");
        
        return headerFooterContext -> {
            return headerFooterContext.isFooter() ? Optional.of(values) : Optional.empty();
        };
    }
    
    /**
     * Anweisung Kopfzeilen im Ergebnis-Dokument mit den übergeben Werten zu ersetzen wenn der
     * Name der Kopf-Zeile mit dem in {@code headerName} übergebenem übereinstimmt.
     * 
     * <p>Der Name der Kopfzeile ist unabhängig der Groß- und Kleinschreibung.</p>
     * 
     * @param headerName    Name Kopfzeile
     * 
     * @param values        Ersetzungs-Werte
     * 
     * @return              Kopfzeilen-Anweisung
     */
    public static HeaderFooterInstruction replaceHeaderWith(String headerName, DataPage values) {
        Objects.requireNonNull(headerName, "headerName");
        Objects.requireNonNull(values, "values");
        
        return headerFooterContext -> {
            if (headerFooterContext.isFooter()) return Optional.empty();
            if (headerName.equalsIgnoreCase(headerFooterContext.getName())) {
                return Optional.of(values);
            } else {
                return Optional.empty();
            }
        };
    }
    
    /**
     * Anweisung Fußzeilen im Ergebnis-Dokument zu ersetzen mit den übergebenen Werten wenn
     * der Name der Fußzeile mit dem in {@code footerName} übereinstimmt.
     * 
     * <p>Der Name der Fußzeile ist unabhängig der Groß- und Kleinschreibung.</p>
     * 
     * @param footerName    Name Fußzeile
     * 
     * @param values        Ersetzungs-Werte
     * 
     * @return              Fußzeilen-Anweisung
     */
    public static HeaderFooterInstruction replaceFooterWith(String footerName, DataPage values) {
        Objects.requireNonNull(footerName, "footerName");
        Objects.requireNonNull(values, "values");
        
        return headerFooterContext -> {
            if (headerFooterContext.isHeader()) return Optional.empty();
            if (footerName.equalsIgnoreCase(headerFooterContext.getName())) {
                return Optional.of(values);
            } else {
                return Optional.empty();
            }
        };
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////

    /**
     * Anweisung einen Document-Interceptor aufzurufen für den übergebenen Dokumenten-Part (Name
     * der XML Datei im Dokument).
     * 
     * @param partName      Name der XML Datei im Dokument (z.B. {@code 'word/document.xml'}.
     * 
     * @param type          Aufruf vor oder nach der Ausfürung des Ersetzungs-Prozesses.
     *                      Bei {@link DocumentInterceptorType#BEFORE_GENERATION} wird das Dokument
     *                      dem Interceptor vor der Ersetzungs übergeben.
     *                      Bei {@link DocumentInterceptorType#AFTER_GENERATION} wird das Dokument
     *                      nach der Ausführung der Ersetzung übergeben.
     * 
     * @param function      Funktion/Listener der aufgerufen wird um ggf. das Dokument überarbeiten
     *                      zu können.
     * 
     * @param moreValues    Ersetzungs-Werte die der {@code function} bei Aufruf mit übergeben
     *                      werden. Angabe von Ersetzungs-Werten ist optional und kann entfallen.
     * 
     * @return              Anweisung den Interceptor entsprechend aufzurufen/auszuführen.
     */
    public static DocumentInterceptor interceptXmlDocumentPart(
            String partName, DocumentInterceptorType type,
            DocumentInterceptorFunction function,
            @Nullable DataValueMap ... moreValues)
    {
        Objects.requireNonNull(partName, "partName");
        Objects.requireNonNull(type, "type");
        Objects.requireNonNull(function, "function");
        
        if (null == moreValues) {
            moreValues = new DataValueMap[0];
        }
        
        return new DefaultDocumentInterceptor(partName, type, function, moreValues);
    }
    
    /**
     * Anweisung einen Document-Interceptor aufzurufen für den übergeben Dokumenten-Part (Name der
     * XML Datei) bevor der Ersetungs-Vorgang stattfindet.
     * 
     * <p>Identisch zu {@link #interceptXmlDocumentPart(String, DocumentInterceptorType, 
     * DocumentInterceptorFunction, DataValueMap...)}, mit der Vorbelegung 
     * {@link DocumentInterceptorType#BEFORE_GENERATION}.</p>
     * 
     * @param partName      Name der XML Datei
     * 
     * @param function      Callback-Funktion
     * 
     * @param moreValues    Optionale Ersetzungs-Werte
     * 
     * @return              Anweisung den Interceptor aufzurufen.
     */
    public static DocumentInterceptor interceptXmlDocumentPartBefore(
            String partName, DocumentInterceptorFunction function, DataValueMap ... moreValues)
    {
        return interceptXmlDocumentPart(partName, BEFORE_GENERATION, function, moreValues);
    }
    
    /**
     * Anweisung einen Document-Interceptor aufzurufen für den übergeben Dokumenten-Part (Name der
     * XML Datei) nach der Ersetungs-Vorgang stattgefunden hat.
     * 
     * <p>Identisch zu {@link #interceptXmlDocumentPart(String, DocumentInterceptorType, 
     * DocumentInterceptorFunction, DataValueMap...)}, mit der Vorbelegung 
     * {@link DocumentInterceptorType#AFTER_GENERATION}.</p>
     * 
     * @param partName      Name der XML Datei
     * 
     * @param function      Callback-Funktion
     * 
     * @param moreValues    Optionale Ersetzungs-Werte
     * 
     * @return              Anweisung den Interceptor aufzurufen.
     */
    public static DocumentInterceptor interceptXmlDocumentPartAfter(
            String partName, DocumentInterceptorFunction function, DataValueMap ... moreValues)
    {
        return interceptXmlDocumentPart(partName, AFTER_GENERATION, function, moreValues);
    }
    
    public static DocumentInterceptor interceptDocumentBody(DocumentInterceptorType type,
            DocumentInterceptorFunction function, DataValueMap ... moreValues)
    {
        return interceptXmlDocumentPart(GENERIC_PART_BODY, type, function, moreValues);
    }
    
    public static DocumentInterceptor interceptDocumentStyles(DocumentInterceptorType type,
            DocumentInterceptorFunction function, DataValueMap ... moreValues)
    {
        return interceptXmlDocumentPart(GENERIC_PART_STYLES, type, function, moreValues);
    }
    
    public static DocumentInterceptor interceptDocumentBodyBefore(
            DocumentInterceptorFunction function, DataValueMap ... moreValues)
    {
        return interceptDocumentBody(BEFORE_GENERATION, function, moreValues);
    }
    
    public static DocumentInterceptor interceptDocumentBodyAfter(
            DocumentInterceptorFunction function, DataValueMap ... moreValues)
    {
        return interceptDocumentBody(AFTER_GENERATION, function, moreValues);
    }
    
    public static DocumentInterceptor interceptDocumentStylesBefore(
            DocumentInterceptorFunction function, DataValueMap ... moreValues)
    {
        return interceptDocumentStyles(BEFORE_GENERATION, function, moreValues);
    }
    
    public static DocumentInterceptor interceptDocumentStylesAfter(
            DocumentInterceptorFunction function, DataValueMap ... moreValues)
    {
        return interceptDocumentStyles(AFTER_GENERATION, function, moreValues);
    }

}
