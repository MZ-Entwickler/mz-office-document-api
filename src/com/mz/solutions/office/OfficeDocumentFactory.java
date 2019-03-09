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
package com.mz.solutions.office;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import static java.util.stream.Collectors.toList;
import java.util.stream.Stream;
import javax.annotation.concurrent.NotThreadSafe;
import static mz.solutions.office.resources.MessageResources.formatMessage;
import static mz.solutions.office.resources.OfficeDocumentFactoryKeys.INVALID_PROP_VALUE;
import static mz.solutions.office.resources.OfficeDocumentFactoryKeys.NO_PROPERTY;

/**
 * Fabrik zum Wechsel von Office-Implementierungen für das Einsetzen in
 * Vorlagen mittels dem Daten-Modell.
 * 
 * @author  Riebe, Moritz   (moritz.riebe@mz-solutions.de)
 */
@NotThreadSafe
public abstract class OfficeDocumentFactory {
    
    ////////////////////////////////////////////////////////////////////////////
    // Factory: Verwaltung aller Office-Implementierungen
    //
    // Jede Office-Implementierung erbt von OfficeDocumentFactory und
    // registriert sich selbstständig per static-initializer.
    //
    // Die Apache OpenOffice und Microsoft Office Implementierungen werden
    // bereits hier manuell eingetragen.
    //
    
    /**
     * Funktionale Schnittstelle für alle Implementierungen die sich als
     * Factory registrieren wollen.
     * 
     * <p>Implementierungs-Hinweis: Alle Implementierungen sollten auf eine
     * schnelle und ressourcen-sparende Instanzierung achten. Für jedes
     * automatische Scannen des Dokument-Types wird eine eigene Factory von
     * jeder Implementierung angelegt.</p>
     */
    @FunctionalInterface
    protected static interface OfficeDocumentFactoryConstructor {
        
        /**
         * Liefert eine Instanz der Factory-Implementierung zurück.
         * 
         * @return  Rückgabe von {@code null} ist nicht zulässig
         */
        public OfficeDocumentFactory createFactoryInstance();
    }
    
    /**
     * Enthält eine Liste aller Klassen (Konstruktoren, welche die API für ein
     * spezielles Office implementieren.
     * 
     * <p>Die Registrierung erfolgt über {@link}</p>
     */
    private static final List<OfficeDocumentFactoryConstructor> factoryImpls;
    
    /**
     * Registriert eine weitere Office-Implementierung.
     * 
     * @param officeFactoryConstructor
     *        Konstruktor-Funktion zur Erstellung einer Factory-Instanz der
     *        entsprechenden Implementierung.
     */
    protected static void registerOfficeImplementation(
            final OfficeDocumentFactoryConstructor officeFactoryConstructor) {

        Objects.requireNonNull(
                officeFactoryConstructor,
                "officeFactoryConstructor");
        
        factoryImpls.add(officeFactoryConstructor);
    }
    
    static {
        factoryImpls = Collections.synchronizedList(new ArrayList<>(4));
        
        // MS Office und OpenOffice werden hier manuell eingetragen!
        registerOfficeImplementation(OpenDocumentFactory::new);
        registerOfficeImplementation(MicrosoftDocumentFactory::new);
    }
    
    ////////////////////////////////////////////////////////////////////////////
    // API zum Nutzen der OfficeDocument API
    //
    // Das Verhalten dieser Methoden sollte bitte nicht verändert werden.
    //
    
    /**
     * Erzeugt eine neue Factory for Apache Open Office und LibreOffice
     * Dokumente.
     * 
     * @return  Factory-Instanz
     */
    public static OfficeDocumentFactory newOpenOfficeInstance() {
        return new OpenDocumentFactory();
    }
    
    /**
     * Erzeugt eine neue Factory für Microsoft Office Dokumente ab
     * Version 2007 im DOCX Format.
     * 
     * @return  Factory-Instanz
     */
    public static OfficeDocumentFactory newMicrosoftOfficeInstance() {
        return new MicrosoftDocumentFactory();
    }
    
    /**
     * Wählt die passende Implementierung anhand des übergebenen Dokumentes aus.
     * 
     * <p>Voraussetzung dafür ist, das das übergebene Dokument mehrfach geöffnet
     * und gelesen werden darf, damit jede Implementierung überprüfen kann
     * ob es sich um ein kompatibles Format handelt.</p>
     * 
     * @param sourceDocument    Dokument für das eine passende Implementierung
     *                          gewählt werden soll; darf nicht
     *                          {@code null} sein.
     * 
     * @return  Factory-Implementierung oder {@code Optional.empty()}.
     */
    public static Optional<OfficeDocumentFactory> newInstanceByDocumentType(
            final Path sourceDocument) throws OfficeDocumentException {

        Objects.requireNonNull(sourceDocument, "sourceDocument");
        
        final List<OfficeDocumentFactory> factories = factoryImpls.stream()
                .map(OfficeDocumentFactoryConstructor::createFactoryInstance)
                .collect(toList());
        
        return factories.stream()
                .filter(factory -> factory.isMyDocumentType(sourceDocument))
                .findFirst();
    }
    
    ////////////////////////////////////////////////////////////////////////////

    private final Map<OfficePropertyType, Object> properties;
    
    {
        properties = new IdentityHashMap<>();
    }
    
    protected OfficeDocumentFactory() {
        setProperty(OfficeProperty.ERR_ON_MISSING_VAL, Boolean.TRUE);
        setProperty(OfficeProperty.ERR_ON_VER_MISMATCH, Boolean.FALSE);
        setProperty(OfficeProperty.ERR_ON_NO_DATA, Boolean.TRUE);
        setProperty(OfficeProperty.IMG_LOAD_AND_EMBED_EXTERNAL, Boolean.TRUE);
    }
    
    /**
     * Setzt eine Einstellung für den Umgang und das Verhalten um Umgang
     * mit Office Dokumenten.
     * 
     * <p>Die in {@link OfficeProperty} vordefinierten Einstellungen werden von
     * jeder Implementierung angeboten. Implementierungsspezifische sind
     * dessen Dokumentation zu entnehmen. Gesetze Einstellungen können auch
     * später überschrieben werden und werden dann neu angewendet auf <u>alle
     * bisher geöffneten {@link OfficeDocument}e und später zu öffnenten</u>
     * Dokumente. Somit wirkt sich eine Einstellungsänderung nicht nur auf
     * nachfolgend geöffnete Dokumente aus!</p>
     * 
     * @param <T>           Typ der Wertes für die Einstellung.
     * 
     * @param property      Instanz der Property-Klasse; z.B.
     *                      {@code OfficeProperty.ERR_ON_VER_MISMATCH}.
     * 
     * @param value         Wert der zur Einstellung gesetzt werden soll.
     */
    public final <T> void setProperty(OfficePropertyType<T> property, T value) {
        Objects.requireNonNull(property, "property");
        
        if (property.isValidPropertyValue(value) == false) {
            throw new IllegalArgumentException(formatMessage(INVALID_PROP_VALUE,
                    /* {0} */ property.name()));
        }
        
        properties.put(property, value);
    }
    
    /**
     * Abrufen einer bereits gesetzten Einstellung.
     * 
     * <p>Es können nur Einstellung abgefragt werden, die bereits zuvor
     * gesetzt wurden. Für jede Option (Standardeinstellungen) und für alle
     * implementierungs-spezifischen Optionen existieren immer 
     * Voreinstellungen mit gültigem Wert.</p>
     * 
     * @param <T>       Typ des Wertes für die Einstellung.
     * 
     * @param property  Instanz der Property-Klasse; z.B.
     *                  {@code OfficeProperty.ERR_ON_VER_MISMATCH}
     * 
     * @return          Wert der Einstellung
     */
    public final <T> T getProperty(OfficePropertyType<T> property) {
        Objects.requireNonNull(property, "property");
        
        final T result = (T) properties.get(property);
        
        if (property.isValidPropertyValue(result) == false) {
            // Dann wurde wahrscheinlich keine Einstellung gesetzt
            throw new IllegalArgumentException(formatMessage(NO_PROPERTY,
                    /* {0} */ property.name()));
        }
        
        return result;
    }
    
    /**
     * Öffnet ein Office-Dokument basierend auf dem Format der Office
     * Implementierung für Ersetzungsvorgänge.
     * 
     * <p>Die Implementierung wirft keine checked Exceptions und verpackt
     * {@link IOException}'s als {@link UncheckedIOException} um die
     * Verwendung in {@link Stream} zu vereinfachen.</p>
     * 
     * @param document  Dokument mit mindestens Leserechten; darf
     *                  nicht {@code null} sein.
     * 
     * @return          Dokument
     * 
     * @throws          UncheckedIOException
     *                  Im Falle eines I/O Fehlers (z.B. beim Öffnen
     *                  des Dokumentes).
     * 
     * @throws          OfficeDocumentException
     *                  Wirft Unterklassen dieser Exception um Fehler beim
     *                  Öffnen oder bei der Erstellung anzuzeigen.
     */
    public abstract OfficeDocument openDocument(Path document)
            throws OfficeDocumentException;
    
    ///// Interne Implementierungen ////////////////////////////////////////////
    
    /**
     * Überprüft anhand des übergebenen Dokumentes ob diese Implementierung
     * dazu passend ist und damit umgehen kann.
     * 
     * <p>Die Implementierung darf keine Exceptions werfen. Auch keine
     * {@link RuntimeException}'s!!</p>
     * 
     * @param document  Dokument das zu prüfen ist; {@code null} wird nicht
     *                  übergeben
     * 
     * @return          bei {@code true} ist die Implementierung mit dem
     *                  übergebenen Dokument kompatibel
     */
    protected abstract boolean isMyDocumentType(Path document);
    
    ////////////////////////////////////////////////////////////////////////////
    // Utility-Methoden zu Vereinfach der Formatüberprüfung
    ////////////////////////////////////////////////////////////////////////////
    
    protected boolean isFileAccessible(Path file) {
        return null != file
                && Files.isDirectory(file) == false
                && Files.isReadable(file);
    }
    
    protected byte[] readDataFromFile(Path file, int maxReadBuffer)
            throws IOException {
        
        // Der folgende Part SOLLTE DRINGEND neu geschrieben werden
        // weil arsch lahm und ziemlich RAM fressend!!!!
        
        final ByteArrayOutputStream byteOut =
                new ByteArrayOutputStream(maxReadBuffer);
        
        try (InputStream inFile = Files.newInputStream(file)) {
            
            int read = 0;
            int lastRead = 0;
            
            while (read < maxReadBuffer && lastRead != -1) {
                lastRead = inFile.read();
                
                if (lastRead != -1) {
                    byteOut.write(lastRead);
                }
                
                read++;
            }
        }
        
        return byteOut.toByteArray();
    }
    
    protected boolean contains(byte[] data, byte[] pattern) {
        final int[] searchPattern = buildSearchPattern(pattern);

        int j = 0;
        if (data.length == 0) {
            return false;
        }

        for (int i = 0; i < data.length; i++) {
            while (j > 0 && pattern[j] != data[i]) {
                j = searchPattern[j - 1];
            }
            
            if (pattern[j] == data[i]) {
                j++;
            }
            
            if (j == pattern.length) {
                return true;
            }
        }
        
        return false;
    }
    
    private int[] buildSearchPattern(byte[] pattern) {
        final int[] searchPattern = new int[pattern.length];
        int j = 0;
        
        for (int i = 1; i < pattern.length; i++) {
            while (j > 0 && pattern[j] != pattern[i]) {
                j = searchPattern[j - 1];
            }
            
            if (pattern[j] == pattern[i]) {
                j++;
            }
            
            searchPattern[i] = j;
        }

        return searchPattern;
    }
    
}
