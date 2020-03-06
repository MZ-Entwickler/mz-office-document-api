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
 * Oberklasse aller Exceptions die im Zusammenhang mit dem Befüllen/ Erstellen
 * und Öffnen von Dokumenten zu tun haben.
 * 
 * <p>Die Klasse wird nicht direkt geworfen; nur spezialisierte Unterklassen
 * davon werden geworfen.</p>
 * 
 * @author  Riebe, Moritz   (moritz.riebe@mz-entwickler.de)
 */
public class OfficeDocumentException extends RuntimeException {

    public OfficeDocumentException(String message) {
        super(message);
    }

    public OfficeDocumentException(String message, Throwable cause) {
        super(message);
        addSuppressed(cause);
    }

    public OfficeDocumentException(Throwable cause) {
        super(cause.getMessage(), cause);
    }

    ////////////////////////////////////////////////////////////////////////////
    
    /**
     * Wenn das Dokument im Format der Office Implementierung ist, jedoch mit
     * einer neueren oder zu alten Office-Version erstellt wurde.
     */
    public static class DocumentVersionMismatchException
            extends OfficeDocumentException {

        public DocumentVersionMismatchException(String message) {
            super(message);
        }

        public DocumentVersionMismatchException(
                String message, Throwable cause) {
            super(message, cause);
        }

        public DocumentVersionMismatchException(Throwable cause) {
            super(cause);
        }
        
    }
    
    /**
     * Für Fehler die während der Dokumenterstellung auftreten.
     */
    public static class FailedDocumentGenerationException
            extends OfficeDocumentException {

        public FailedDocumentGenerationException(String message) {
            super(message);
        }

        public FailedDocumentGenerationException(
                String message, Throwable cause) {
            super(message, cause);
        }

        public FailedDocumentGenerationException(Throwable cause) {
            super(cause);
        }
        
    }
    
    /**
     * Zu einem Platzhalter im Dokument existiert kein entsprechender Wert.
     */
    public static class DocumentPlaceholderMissingException
            extends FailedDocumentGenerationException {

        public DocumentPlaceholderMissingException(String message) {
            super(message);
        }

        public DocumentPlaceholderMissingException(
                String message, Throwable cause) {
            super(message, cause);
        }

        public DocumentPlaceholderMissingException(Throwable cause) {
            super(cause);
        }
        
    }
    
    /**
     * Das übergebene Format entspricht nicht dem erwarteten Format.
     */
    public static class InvalidDocumentFormatForImplementation
            extends OfficeDocumentException {

        public InvalidDocumentFormatForImplementation(String message) {
            super(message);
        }

        public InvalidDocumentFormatForImplementation(
                String message, Throwable cause) {
            super(message, cause);
        }

        public InvalidDocumentFormatForImplementation(Throwable cause) {
            super(cause);
        }
        
    }
    
    /**
     * Zur Dokumenterstellung liegen keine Daten vor.
     * 
     * <p>Tritt nur ein, sobald die Übergebene Menge an {@link DataPage}'es
     * gleich 0 ist und somit keine Werte vorliegen.</p>
     */
    public static class NoDataForDocumentGenerationException
            extends FailedDocumentGenerationException {

        public NoDataForDocumentGenerationException(String message) {
            super(message);
        }

        public NoDataForDocumentGenerationException(
                String message, Throwable cause) {
            super(message, cause);
        }

        public NoDataForDocumentGenerationException(Throwable cause) {
            super(cause);
        }
        
    }
    
}
