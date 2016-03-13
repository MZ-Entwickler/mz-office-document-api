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
package com.mz.solutions.office.model;

import java.io.Serializable;
import java.util.Arrays;
import java.util.Collections;
import java.util.EnumSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.StringJoiner;
import static mz.solutions.office.resources.DataValueKeys.CONTAINS_SPACES;
import static mz.solutions.office.resources.DataValueKeys.KEY_TOO_SHORT;
import static mz.solutions.office.resources.DataValueKeys.OPTIONS_IN_CONFLICT;
import static mz.solutions.office.resources.DataValueKeys.TOO_MANY_VALUE_OPTIONS;
import static mz.solutions.office.resources.DataValueKeys.UNKNOWN_OPTION_STATE;
import static mz.solutions.office.resources.MessageResources.formatMessage;

/**
 * Einzelner Eintrag für einen Ersetzungsvorgang aus Bezeichner
 * (Platzhaltername) und Wert (Text).
 * 
 * <p>Ein Objekt vom Typ {@code DataValue} ist effektiv nach der Erstellung
 * über den Konstruktor nicht mehr veränderbar.<br><br>
 * 
 * Als Standard-Optionen ist das beibehalten von Zeilenumbrüchen und
 * Tabulator-Zeichen eingerichtet.<br><br>
 * 
 * <b>Platzhalter-Bezeichner</b><br>
 * Der Platzhalter(bezeichner) ist unabhängig von Groß- und Kleinschreibung,
 * muss aus mindestens 2 Zeichen bestehen und darf keine Leerzeichen enthalten.
 * Um Unterschiede zwischen Office-Implementierungen zu vermeiden, sollte der
 * Zeichenvorrat auf {@code [A-Z, a-z, 0-9, '_']} beschränkt bleiben.</p>
 * 
 * @author  Riebe, Moritz       (moritz.riebe@mz-solutions.de)
 */
public final class DataValue implements Serializable {
    
    private static final ValueOptions DEFAULT_OPTION_TABULATOR;
    private static final ValueOptions DEFAULT_OPTION_LINEBREAK;
    
    private static final Set<ValueOptions> defaultOptions;
    
    static {
        // Bei Veränderung der Vorbelegung siehe #prepareValueByOptions(..) !
        // Sowie Überprüfung auf Konflikt in #prepareValueOptions(..) !
        //
        // Eine Veränderung der Vorbelegung sollte auch in der JavaDoc
        // Dokumentation von DataValue (Klassen-Doc) und ValueOptions
        // vermerkt werden.
        
        DEFAULT_OPTION_LINEBREAK = ValueOptions.KEEP_LINEBREAK;
        DEFAULT_OPTION_TABULATOR = ValueOptions.KEEP_TABULATOR;
        
        final Set<ValueOptions> modifiableOptions = EnumSet.of(
                DEFAULT_OPTION_LINEBREAK,
                DEFAULT_OPTION_TABULATOR);
        
        defaultOptions = Collections.unmodifiableSet(modifiableOptions);
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private final Set<ValueOptions> options;
    
    private final String keyName;
    private final String value;
    
    /**
     * Erzeugt ein Platzhalter Bezeichner-Werte-Paar mit Vorsteinstellungen.
     * 
     * <p>Genauere Beschreibung findet sich in der Konstruktorbeschreibung von:
     * {@link #DataValue(String, String, ValueOptions...)}. </p>
     * 
     * @param keyName   Bezeichner des Platzhalters; Groß- und Kleinschreibung
     *                  ist unrelevant; mindestens zwei Zeichen; darf nicht
     *                  {@code null} sein und darf keine Leerzeichen enthalten
     * 
     * @param value     Wert des Platzhalters; bei {@code null} wird von einer
     *                  leeren Zeichenkette ausgegangen
     */
    public DataValue(String keyName, String value) {
        this.keyName = prepareKeyName(keyName);
        
        this.options = defaultOptions;
        this.value = prepareValueByOptions(value);
    }
    
    /**
     * Erzeugt ein Platzhalter Bezeichner-Werte-Paar.
     * 
     * <p>Bei Angabe der Optionen darf jeweils nur eine Option für Tabulatoren
     * und eine für Zeilenumbrüche gewählt werden, damit diese nicht im
     * Widerspruch zueinander stehen. Wird nur eine Option angeben, wird die
     * fehlende mit der Voreinstellung belegt.</p>
     * 
     * @param keyName       Bezeichner des Platzhalters; Groß- und
     *                      Kleinschreibung ist unrelevant; mindestens zwei
     *                      Zeichen; darf nicht {@code null} sein und darf
     *                      keine Leerzeichen enthalten
     * 
     * @param value         Wert des Platzhalters; bei {@code null} wird von
     *                      einer leeren Zeichenkette ausgegangen
     * 
     * @param valueOptions  Angabe von Optionen die auf den Wert angewendetet
     *                      werden sollen beim Ersetzungsvorgang
     */
    public DataValue(String keyName, String value,
            ValueOptions ... valueOptions) {
        
        this.keyName = prepareKeyName(keyName);
        
        this.options = prepareValueOptions(valueOptions);
        this.value = prepareValueByOptions(value);
    }
    
    private String prepareKeyName(String keyName) {
        Objects.requireNonNull(keyName, "keyName");
        
        final String preparedKeyName = keyName.trim().toUpperCase();
        
        if (preparedKeyName.length() <= 1) {
            throw new IllegalArgumentException(formatMessage(KEY_TOO_SHORT, preparedKeyName));
        }
        
        if (preparedKeyName.contains(" ")) {
            throw new IllegalArgumentException(formatMessage(CONTAINS_SPACES, preparedKeyName));
        }
        
        return preparedKeyName;
    }
    
    private Set<ValueOptions> prepareValueOptions(
            ValueOptions ... valueOptions) {
        
        if (null == valueOptions || valueOptions.length == 0) {
            return defaultOptions;
        }
        
        final List<ValueOptions> optionList = Arrays.asList(valueOptions);
        final Set<ValueOptions> optionSet = EnumSet.copyOf(optionList);
        
        if (optionSet.size() > 2) {
            throw new IllegalArgumentException(formatMessage(TOO_MANY_VALUE_OPTIONS, optionSet));
        }
        
        final Optional<ValueOptions> lineBreakOption = optionSet.stream()
                .filter(ValueOptions::isLineBreakOption)
                .findFirst();
        
        final Optional<ValueOptions> tabulatorOption = optionSet.stream()
                .filter(ValueOptions::isTabulatorOption)
                .findFirst();
        
        // Verhindern von widersprüchlichen Optionen
        // Gibt es insgesamt zwei Optionen, muss EINE für den Tabulator sein
        // und die ANDERE für Zeilenumbrüche
        
        if (optionSet.size() == 2) {
            if (lineBreakOption.isPresent() == false
                    || tabulatorOption.isPresent() == false) {
                throw new IllegalArgumentException(formatMessage(OPTIONS_IN_CONFLICT, optionSet));
            }
                    
            
            final Set<ValueOptions> resultSet = EnumSet.of(
                    lineBreakOption.get(),
                    tabulatorOption.get());
            
            return Collections.unmodifiableSet(resultSet);
        }
        
        // Status:  nur eine von beiden Optionen gewählt (Tab oder Linebreak);
        //          belegen der fehlenden Option mit der Voreinstellung
        
        assert tabulatorOption.isPresent() || lineBreakOption.isPresent();
        
        final Set<ValueOptions> resultSet = EnumSet.of(
                tabulatorOption.orElse(DEFAULT_OPTION_TABULATOR),
                lineBreakOption.orElse(DEFAULT_OPTION_LINEBREAK));
        
        return Collections.unmodifiableSet(resultSet);
    }
    
    private String prepareValueByOptions(String value) {
        // Aufruf setzt voraus, dass this.options bereits korrekt belegt ist!
        
        String pValue = null == value ? "" : value;
        
        if (pValue.isEmpty() || options == defaultOptions) {
            // Wenn der Wert leer ist, bringt die Durchführung der Optionen
            // absolut nichts
            
            // Die vorbelegten Standardoptionen erzeugen keine Änderungen
            // somit kann auch vorzeitig abgebrochen werden
            return pValue;
        }
        
        for (ValueOptions singleOption : options) {
            switch (singleOption) {
                case LINEBREAK_TO_WHITESPACE:
                    pValue = pValue.replace('\n', ' ');
                    pValue = pValue.replace("\r", "");
                    break;
                    
                case TABULATOR_TO_FOUR_SPACES:
                    pValue = pValue.replace("\t", "    ");
                    break;
                    
                case TABULATOR_TO_WHITESPACE:
                    pValue = pValue.replace('\t', ' ');
                    break;
                    
                case KEEP_LINEBREAK:
                case KEEP_TABULATOR:
                    // Konvertierung erfolgt nicht in der Zeichenkette sondern
                    // später im Formatierungsbaum
                    break;
                    
                default:
                    throw new IllegalStateException(formatMessage(UNKNOWN_OPTION_STATE,
                            /* {0} */ ValueOptions.class.getSimpleName(),
                            /* {1} */ singleOption.name()));
            }
        }
        
        return pValue;
    }
    
    /**
     * Rückgabe des Platzhalterbezeichners in Großbuchstaben.
     * 
     * @return      Platzhalterbezeichner
     */
    public String getKeyName() {
        return this.keyName;
    }
    
    /**
     * Rückgabe des Ersetzungswertes/ Platzhalterwert.
     * 
     * <p>Der zurückgebene Wert muss nicht unbedingt identisch sein mit dem
     * im Konstruktor angegebenen; durch die Optionen können bereits
     * Tabulatoren oder Zeilenumbrüche umgewandelt worden sein.</p>
     * 
     * @return      Platzhalterwert; nie {@code null}.
     */
    public String getValue() {
        return this.value;
    }
    
    /**
     * Angegebene Optionen bei der Einsetzung wie der Platzhalterwert
     * behandelt werden soll.
     * 
     * @return      Gewählte Optionen; Menge ist nicht veränderbar
     */
    public Set<ValueOptions> getValueOptions() {
        return this.options;
    }

    /**
     * Erzeugt einen möglichst gleich verteilten Hash-Code.
     * 
     * <p>Bei der Generierung des Hash-Codes wird nicht zwischen Groß-
     * und Kleinschreibung unterschieden.</p>
     * 
     * @return      Hash-Code
     */
    @Override
    public int hashCode() {
        return Objects.hashCode(keyName);
    }

    /**
     * Vergleicht dieses Objekt mit dem übergebenen.
     * 
     * <p>Beide Objekte sind nur gleich, wenn das übergebene Objekt nicht
     * {@code null} ist, vom selben Typ und die Platzhalterbezeichner
     * unabhängig der Groß- und Kleinschreibung identisch sind.</p>
     * 
     * <pre>
     *  DataValue a = new DataValue("key", "value");
     *  DataValue b = new DataValue("KEY", "value");
     *  DataValue c = new DataValue("key", "other value");
     *  DataValue d = new DataValue("yek", "12345");
     * 
     *  x.equals(y)|    a       b       c       d
     *  -----------+---------------------------------------
     *      a      |  true    true    true    false
     *      b      |  true    true    true    false
     *      c      |  true    true    true    false
     *      d      |  false   false   false   true
     * 
     * </pre>
     * 
     * @param obj   Vergleichsobjekt
     * 
     * @return      {@code true}, wenn beide {@code DataValue}'s den selben
     *              Platzhalterbezeichner besitzen
     */
    @Override
    public boolean equals(Object obj) {
        if (obj == null) {
            return false;
        }
        
        if (obj == this) {
            return true;
        }
        
        if (!(obj instanceof DataValue)) {
            return false;
        }
        
        final DataValue other = (DataValue) obj;
        return this.keyName.equals(other.keyName);
    }
    
    /**
     * Wandelt den Platzhalterbezeichner und Platzhalterwert dieses Objektes
     * als zusammenhängende Zeichenkette um mit Angabe der Optionen soweit
     * diese von der Standardbelegung abweichen.
     * 
     * <p>Für ein mit {@code new DataValue("KLNAME", "Meier")} erstelltes
     * Objekt wäre folgende Rückgabe:</p>
     * <pre>
     *  DataValue: keyName='KLNAME' value='Meier' options=default
     * </pre>
     * 
     * <p>Sobald Optionen von der Standardbelegung abweichen, werden diese
     * einzel aufgezähl.</p><pre>
     *  new DataValue("KLNAME", "Meier",
     *          ValueOptions.KEEP_TABULATOR,
     *          ValueOptions.LINEBREAK_TO_WHITESPACE);
     * 
     *  Erzeigt bei Aufruf von #toString() folgende Rückgabe:
     *  DataValue: keyName='KLNAME' value='Meier' options=[KEEP_TABULATOR, LINEBREAK_TO_WHITESPACE]
     * </pre>
     * 
     * @return      Repräsentation dies Objektes als Zeichenkette; gibt
     *              niemals {@code null} zurück.
     */
    @Override
    public String toString() {
        final StringBuilder builder = new StringBuilder();
        
        builder.append("DataValue: keyName=\'").append(keyName);
        builder.append("\'  value=\'").append(value).append("\' options=");
        
        if (options == defaultOptions) {
            builder.append("default");
        } else {
            StringJoiner joinOptions = new StringJoiner(", ", "[", "]");
            for (ValueOptions singleOption : options) {
                joinOptions.add(singleOption.name());
            }
            
            builder.append(joinOptions.toString());
        }
        
        return builder.toString();
    }
    
}
