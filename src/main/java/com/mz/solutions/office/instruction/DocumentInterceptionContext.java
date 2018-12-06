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

import com.mz.solutions.office.model.DataValueMap;
import org.w3c.dom.Document;
import org.w3c.dom.Node;

import java.util.List;

/**
 * Kontext beim Callback während ein Document-Interceptor aufgerufen wird und seine eigenen
 * Aktionen ausführen kann.
 *
 * @author Riebe, Moritz      (moritz.riebe@mz-solutions.de)
 */
public interface DocumentInterceptionContext extends InstructionContext {

    /**
     * Rückgabe des ursprünglich übergebenen Interceptors.
     *
     * @return Instanz der als {@link DocumentProcessingInstruction} übergebenen Anweisung.
     */
    public DocumentInterceptor getInterceptor();

    /**
     * Name des Abschnittes auf dem Zugriff gewährt wurde.
     *
     * <p>Der Abschnitt kann ggf. {@code "#body"} oder {@code "#styles"} lauten und entspricht
     * dann der jeweiligen Office-Implementierung für den Inhalt und die Formatvorlagen.</p>
     *
     * @return Name des Abschnits im Dokument der jetzt zur Bearbeitung ansteht,
     * z.B. {@code '#body', '#styles', 'META-INF/manifest'}.
     */
    public String getPartName();

    /**
     * Überprüft ob es sich um ein XML basiertes Dokument handelt, bisher wird immer {@code true}
     * zurück gegeben.
     *
     * @return {@code true}.
     */
    public boolean isXmlBasedDocumentPart();

    /**
     * Rückgabe des Document-Interceptor Typs ob dieser <u>vor</u> oder <u>nach</u> dem
     * Generierungs-Prozesses aufgerufen werden soll.
     *
     * @return Instanz von {@link DocumentInterceptorType}.
     */
    public DocumentInterceptorType getInterceptorType();

    /**
     * Rückgabe aller Daten die zur Ersetzung des Dokumentes angewendet werden und einer der
     * {@code generate*} Methoden übergeben wurden.
     *
     * @return Alle Daten die auf dieses Dokument angewendet werden, oder angewendet wurden.
     */
    public List<DataValueMap<?>> getDocumentValues();

    /**
     * Rückgabe von Daten die speziell diesem Interceptor mit übergeben worden sind.
     *
     * @return Menge der Daten.
     */
    public List<DataValueMap<?>> getInterceptorValues();

    /**
     * Rückgabe des betreffenden XML Dokumentes, daran gemachte Änderungen finden sich im
     * Ergebnis-Dokument direkt wieder.
     *
     * @return XML DOM Dokument
     */
    public Document getXmlDocument();

    /**
     * Rückgabe des ersten bearbeitbaren Elementes im XML DOM Dokument.
     *
     * @return XML Node (meistens Element)
     */
    public Node getXmlBodyNode();

}
