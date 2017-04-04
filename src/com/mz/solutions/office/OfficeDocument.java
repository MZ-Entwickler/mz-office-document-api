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
package com.mz.solutions.office;

import com.mz.solutions.office.extension.Extension;
import com.mz.solutions.office.extension.MicrosoftCustomXml;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.result.Result;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.util.Collections;
import java.util.Iterator;
import java.util.Objects;
import java.util.Optional;
import javax.annotation.OverridingMethodsMustInvokeSuper;

/**
 * Oberklasse für Dokumente mit Platzhaltern die mit einem Daten-Modell
 * befüllt werden soll.
 * 
 * <p>Die Klasse kann nur über {@link OfficeDocumentFactory} erstellt werden.
 * Das Verhalten ist bei allen Implementierungen identisch. Das zugrunde
 * liegende Dokument muss jedoch der Implementierung entsprechend angepasst
 * sein.<br>Das geladene Dokument wird im RAM behalten und verarbeitet,
 * ein schließen ist nicht notwendig. Jede Instanz von {@link OfficeDocument}
 * besitzt eine Referenz zur eigenen Factory. Es ist nicht davon auszugehen
 * das die Implementierung thread-safe ist.</p>
 * 
 * <p><b>Erweiterungen:</b> Erweiterungen die abhängig vom Dokument/ Format oder
 * dem Office sind, können spezifisch abgefragt und verwendet werden über die
 * {@link #extension(Class)} Methode.</p>
 * 
 * @author  Riebe, Moritz   (moritz.riebe@mz-solutions.de)
 */
public abstract class OfficeDocument {
    
    OfficeDocument() { }
    
    /**
     * Gibt die Factory zurück mit der dieses Dokument erstellt/ geöffnet wurde.
     * 
     * <p>Geänderte Einstellungen der Factory wirken sich auch auf bereits
     * geöffnete Dokumente aus.</p>
     * 
     * @return      Zugehörige Factory-Implementierung
     */
    public abstract OfficeDocumentFactory getRelatedFactory();
    
    /**
     * Intern - Methode zur Befüllung dieses Dokumentes.
     * 
     * <p>Bei der Implementierung sollte darauf geachtet werden, dass alles
     * soweit möglich im RAM geschieht; keine Ressourcen vergessen werden
     * frei zugeben und diese Methode mehrfach aufgerufen werden kann mit
     * wechselndem Daten-Modell. Die Implementierung muss nicht thread-safe
     * sein und sollte keine Seiteneffekte erzeugen.</p>
     * 
     * @param dataPages     Daten-Modell-Iterator
     * 
     * @return              Erzeugtes Dokument als {@code byte[]}-Array
     * 
     * @throws  IOException
     *          Darf frei geworfen werden.
     * 
     * @throws  OfficeDocumentException
     *          Für Fehler die nicht IO basiert sind, sondern entweder im
     *          Konflikt mit der Konfiguration stehen oder dem Daten-Modell.
     */
    protected abstract byte[] generate(Iterator<DataPage> dataPages)
            throws IOException, OfficeDocumentException;

    /**
     * Ersetzt alle Platzhalter in diesem Dokument und übergibt dies der
     * Ausgabeimplementierung.
     * 
     * <p>Alle übergebenen {@link DataPage}'s werden in dieses Dokument
     * eingesetzt bis der Iterator am Ende ist. Bei einem leeren Iterator ist
     * das Verhalten von der Konfiguration der Factory abhängig. Der Zustand
     * {@code null} wird bei keinem Argument angenommen. Ob mehrere
     * {@link DataPage}'s auch zu mehreren Seiten im Dokument führen ist
     * abhängig der Formatierung des Dokumentes und der Konfiguration der
     * Factory.</p>
     * 
     * @param dataPages     Menge aller Datensätze
     * 
     * @param docResult     Ausgabeimplementierung der das fertig generierte
     *                      Dokument übergeben wird
     * 
     * @throws  UncheckedIOException
     *          Sollte beim Schreibvorgang durch {@code docResult} ein IO Fehler
     *          auftreten, wird dies als {@code unchecked} Exception
     *          weiter gereicht.
     * 
     * @throws  OfficeDocumentException 
     *          Tritt auf wenn Fehler beim Ersetzungsvorgang oder der Erstellung
     *          nicht anhand der Konfiguration behoben werden können.
     */
    public void generate(Iterator<DataPage> dataPages, Result docResult)
            throws UncheckedIOException, OfficeDocumentException {

        Objects.requireNonNull(docResult, "docResult");
        Objects.requireNonNull(dataPages, "dataPages");
        
        try {
            docResult.writeResult(generate(dataPages));
            
        } catch (IOException ioe) {
            throw new UncheckedIOException(ioe);
        }
    }
    
    /**
     * Ersetzt alle Platzhalter in diesem Dokument und übergibt dies der
     * Ausgabeimplementierung.
     * 
     * <p>Siehe {@link #generate(Iterator, Result)}</p>
     * 
     * @param dataPages     Menge aller Datensätze
     * 
     * @param docResult     Ausgabeimplementierung
     */
    public void generate(Iterable<DataPage> dataPages, Result docResult)
            throws UncheckedIOException, OfficeDocumentException {

        Objects.requireNonNull(dataPages, "dataPages");
        
        generate(dataPages.iterator(), docResult);
    }
    
    /**
     * Ersetzt alle Platzhalter in diesem Dokument und übergibt dies der
     * Ausgabeimplementierung.
     * 
     * <p>Siehe {@link #generate(Iterator, Result)}</p>
     * 
     * @param dataPage      Menge aller Datensätze
     * 
     * @param docResult     Ausgabeimplementierung
     */
    public void generate(DataPage dataPage, Result docResult)
            throws UncheckedIOException, OfficeDocumentException {

        Objects.requireNonNull(dataPage, "dataPage");
        
        generate(Collections.singleton(dataPage).iterator(), docResult);
    }
    
    /**
     * Generiert ein Dokument und übergibt dies der Ausgabeimplementierung, ohne
     * ein vorheriges Daten-Modell für den Ersetzungsvorgang anzuwenden.
     * 
     * <p>Diese Methode verhält sich identisch zu {@link #generate(Iterator)}
     * mit dem Unterschied, das kein Daten-Modell übergeben wird. Werden
     * Erweiterungen von diesem Dokument genutzt, dann sollte diese Methode zur
     * Generierung der Nachricht verwendet werden. Ob die verwendete Erweiterung
     * mit einer Übergabe eines Daten-Modells gleichzeitig verwendet werden kann
     * ist unbestimmt.</p>
     * 
     * @param docResult     Ausgabeimplementierung
     */
    public void generate(Result docResult)
            throws UncheckedIOException, OfficeDocumentException {
        
        generate(new DataPage(), docResult);
    }
    
    /**
     * Ermittelt ob dieses Dokument die übergebene Erweiterung unterstützt und
     * gibt bei Unterstützung die Schnittstelle für den Zugriff darauf zurück.
     * 
     * <p>Mehrfaches aufrufen mit der selben Erweiterungsangabe führt immer zur
     * Rückgabe der gleichen Instanz der Erweiterungsimplementierung. Die
     * Art der Erweiterung wird als Typ übergeben.</p>
     * 
     * <pre>
     *  final Optional&lt;MicrosoftCustomXml&gt; opCustXmlExt =
     *          document.extension(MicrosoftCustomXml.class);
     * 
     *  if (opCustXmlExt.isPresent() == false) {
     *      throw new IllegalStateException("Erweiterung wird von diesem "
     *              + "Dokument nicht unterstützt.");
     *  }
     * 
     *  final MicrosoftCustomXml extCustXml = opCustXmlExt.get();
     * 
     *  // Nutzungshinweis: bei mehrmaligem Abrufen/Aufrufen der Erweiterung
     *  // wird NICHT immer eine neue Instanz zurück geliefert.
     *  // Entsprechend ist folgende Aussage korrekt:
     * 
     *  final MicrosoftCustomXml extInstance1 =
     *          document.extension(MicrosoftCustomXml.class).get();
     * 
     *  final MicrosoftCustomXml extInstance2 =
     *          document.extension(MicrosoftCustomXml.class).get();
     * 
     *  // entsprechend gilt:
     * 
     *  (extInstance1 == extInstance2) == true
     *  // Es wird immer die selbe Instanz zurückgeliefert
     * </pre>
     * 
     * @param <T>       Typ der Erweiterung
     * 
     * @param extType   Typ der Erweiterung als Klassen-Typ,
     *                  z.B. {@code MicrosoftCustomXml.class}. Der übergebene
     *                  Typ darf nicht {@code null} sein.
     * 
     * @return          Rückgabe der Erweiterung oder {@code Optional.empty()}
     *                  sobald die Erweiterung nicht verfügbar oder nicht
     *                  implementiert ist.
     * 
     * @see MicrosoftCustomXml
     *      [Erweiterung] XML Mapping für Word-Dokumente ab 2007
     * 
     * @see #generate(com.mz.solutions.office.result.Result)
     *      [Ausgabe] Schreiben von Dokumenten ohne Übergabe eines Daten-Modells
     *                soweit nur die Erweiterung(en) verwendet werden sollen.
     */
    @OverridingMethodsMustInvokeSuper
    public <T extends Extension> Optional<T>  extension(Class<T> extType) {
        Objects.requireNonNull(extType, "extType");
        
        // Implementierende Klassen sollten immer diese Eltern-Klasse-Methode
        // aufrufen bevor die eigene Implementierung die spezifischen Elemente
        // hinzufügt. (Auch wenn dies derzeit keinen höheren Grund hat!)
        //
        // Wenn super.extension(..) aufgerufen wird UND EIN ERGEBNIS LIEFERT
        // MUSS jenes Ergebnis auch zurückgegeben werden.
        
        return Optional.empty();
    }
    
}
