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
package com.mz.solutions.office.model.interceptor;

import com.mz.solutions.office.OfficeDocument;
import com.mz.solutions.office.OfficeDocumentFactory;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataTable;
import com.mz.solutions.office.model.DataTableRow;
import com.mz.solutions.office.model.DataValueMap;
import org.w3c.dom.Document;
import org.w3c.dom.Node;

/**
 * Kontext zur Unterstützung eines {@link ValueInterceptor}'s bei der Erzeugung eines im Kontext
 * passenden Ersetzungswertes.
 * 
 * <p>Der Kontext darf nicht zwischengespeichert werden und ist nur innerhalb der Methode (im selben
 * Thread), als Thread-safe zu betrachten.</p>
 * 
 * <p>Minimum unterstützted Profil; Pflichtmethoden die immer verfügbar sein müssen:<br>
 * {@link #getDocumentFactory()}, {@link #getDocument()}, {@link #getPlaceholderName()} und
 * {@link #getParentValueMap()}.</p>
 * 
 * <p>Andere Methoden sind vom Grad der Implementierung abhängig durch die Office Implementierung
 * und können mit {@link #hasLowLevelAccess()} und {@link #hasXmlSupport()} auf Unterstützung
 * abgefragt werden.</p>
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-solutions.de)
 */
public abstract class InterceptionContext {
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // METHODEN DIE VERPFLICHTENT UNTERSTÜTZT WERDEN MÜSSEN
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    /**
     * Liefert die Factory die zum aktuell ersetzten Dokument zurück und entspricht somit der
     * Office-Implementierung.
     * 
     * @return      Office-Document-Factory zu gerade zu ersetzenden Dokument.
     */
    public abstract OfficeDocumentFactory getDocumentFactory();
    
    /**
     * Aktuelles Dokument (Vorlage) welches ersetzt wird; eventuelle Abfragen auf Unterstützte
     * Erweiterungen sind von diesem Dokument zu erfragen und zu erstellen.
     * 
     * @return      Dokument im Ersetzungsprozess (die entsprechende Vorlage).
     */
    public abstract OfficeDocument getDocument();
    
    /**
     * Name des Platzhalters für den der Ersetzungswert abgefragt/ erzeugt werden soll; dieser
     * Name ist auch im Ergebnis des Interceptors wieder anzugeben.
     * 
     * @return      Name des Platzhalters der ersetzt wird/ werden soll.
     */
    public abstract String getPlaceholderName();
    
    /**
     * Alle Platzhalter im aktuellen Scope und an der aktuellen Stelle des zu erzeugenden
     * Ersetzungswertes.
     * 
     * <p>Zurück gegeben können z.B {@link DataPage}, {@link DataTable}, {@link DataTableRow} welche
     * an der zu ersetzten Stelle jeweils dem Kontext/Scope entsprechen.</p>
     * 
     * <p>Änderungen in der zurückgegeben Map sind zulässig und können ggf. nachfolgend zu
     * ersetzende Platzhalter damit beeinflussen.</p>
     * 
     * @return      Alle Platzhalter (Name und Werte) im Scope des aktuell zu ersetzenden
     *              Platzhalters.
     */
    public abstract DataValueMap<?> getParentValueMap();
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // OPTIONAL GEWÄHRT DIE IMPLEMENTIERUNG AUCH EINEN LOW-LEVEL ZUGRIFF UM DEN/DIE PLATZHALTER
    // SELBER AUFZULÖSEN UND GGF. WEITERREICHENDE ÄNDERUNGEN VORZUNEHMEN
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    /**
     * Prüft ob die Office-Implementierung den Low-Level Zugriff auf die Dokumentstruktur zulässt
     * und unterstützt.
     * 
     * <p>Der entsprechende Handler und Zugriff auf die Low-Level-Funktionen ist dabei völlig
     * abhängig von der entsprechenden Implementierung. Der Handler zum zugriff kann mit
     * {@link #getLowLevelHandler()} abgerufen werden, wenn diese Funktion {@code true} liefert.</p>
     * 
     * @return      {@code true} für Unterstützung der manuellen Bearbeitung.
     */
    public boolean hasLowLevelAccess() {
        return false;
    }
    
    /**
     * Rückgabe des Handlers für Low-Level Zugriff auf das zu ersetzende Dokument.
     * 
     * <p>Vor Aufruf dieser Funktion, sollte die Unterstützung mit {@link #hasLowLevelAccess()}
     * geprüft werden!</p>
     * 
     * @return      Handler
     * 
     * @throws      ValueInterceptorException 
     *              Wenn keine Low-Level Unterstützung vorliegt.
     * 
     * @see         #hasLowLevelAccess()
     */
    public Object getLowLevelHandler() throws ValueInterceptorException {
        throw new ValueInterceptorException.NoLowLevelSupportException("");
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // UNTERSTÜTZUNG FÜR XML BASIERTE DOKUMENTE (MS Office, LibreOffice, XHTML, ...)
    ////////////////////////////////////////////////////////////////////////////////////////////////

    /**
     * Prüft ob das zu ersetzende Dokument in irgendeiner Form XML basiert ist - gibt aber keinen
     * Rückschluss auf die Unterstützung von direktem XML Zugriff.
     * 
     * @return      {@code true}, wenn das Dokument in irgendeiner Form XML basiert ist.
     */
    public boolean isXmlBasedDocument() {
        return false;
    }
    
    /**
     * Prüft auf Unterstützung von direktem (manuellen) XML Zugriff auf das darunter liegende
     * Dokument.
     * 
     * @return      {@code true}, wenn direkter/manueller XML Zugriff auf das Dokument
     *              unterstützt wird
     */
    public boolean hasXmlSupport() {
        return false;
    }
    
    /**
     * Rückgabe des Dokumentes vom XML DOM - nur bei Unterstützung.
     * 
     * <p>XML Unterstützung zuvor mit {@link #hasXmlSupport()} überprüfen!</p>
     * 
     * @return      XML DOM Dokument
     * 
     * @throws      ValueInterceptorException 
     *              Wenn keine Unterstützung von XML Zugriff implementiert ist.
     */
    public Document getXmlDocument() throws ValueInterceptorException {
        throw new ValueInterceptorException.NoXmlBasedDocumentException("");
    }
    
    /**
     * Rückgabe des XML DOM Nodes an dem der aktuelle Platzhalter gefunden wurde.
     * 
     * <p>XML Unterstützung zuvor mit {@link #hasXmlSupport()} überprüfen!</p>
     * 
     * @return      XML DOM Node an der Stelle des Platzhalters
     * 
     * @throws      ValueInterceptorException 
     *              Wenn keine Unterstützung von XML Zugriff implementiert ist.
     */
    public Node getXmlCurrentNode() throws ValueInterceptorException {
        throw new ValueInterceptorException.NoXmlBasedDocumentException("");
    }
    
}
