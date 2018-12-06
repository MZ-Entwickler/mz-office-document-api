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
package com.mz.solutions.office.util;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import javax.annotation.concurrent.NotThreadSafe;
import java.text.CharacterIterator;
import java.text.StringCharacterIterator;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.*;
import java.util.function.Supplier;
import java.util.logging.Logger;

import static com.mz.solutions.office.util.WindowsFileName.MAX_NAME_LENGTH;
import static com.mz.solutions.office.util.WindowsFileName.isLetterForbidden;
import static java.text.CharacterIterator.DONE;
import static java.util.Objects.requireNonNull;

/**
 * Helfer-Klasse zum Einsetzen von Platzhaltern in Zeichenketten für generelle
 * Fälle und zur Erzeugung von Dateinamen.
 * <br><br>
 *
 * <b>Beschreibung</b><br>
 * Ermöglicht die Einsetzung mehrerer frei definierter Platzhalter und einigen
 * standardmäßig vordefinierten in Zeichenketten.
 * <br><br>
 * <p>
 * Vordefinierte Platzhalter sind unabhängig der Verwendeten {@code Map} immer
 * verfügbar und liegen in Deutsch und Englisch vor.<br><br>
 * <p>
 * Um eine ständige Neugenerierung der vordefinierten Platzhalter zu vermeiden
 * sollte eine erzeugte Instanz von {@code PlaceMarkerInserter} mehrfach
 * verwendet werden, soweit die möglich ist.<br><br>
 *
 * <b>Vordefinierte Platzhalter:</b><br>
 * <pre>
 * sys.year        sys.jahr           := Jahreszahl  (z.B. 2014)
 * sys.month       sys.monat          := Monatszahl  (z.B. 6)
 * sys.month.name  sys.monat.name     := Monatsname  (z.B. Juni o. June)
 * sys.day         sys.tag            := Tag im Monat(z.B. 8)
 * sys.weekday     sys.wochentag      := Wochentag   (z.B. Montag o. Monday)
 * sys.date.iso    sys.datum.iso      := Datum       (z.B. 2014-06-08)
 * sys.date.din    sys.datum.din      := Datum       (z.B. 08.06.2014)
 * </pre>
 *
 * <b>Verwendung als Dateiname</b><br>
 * Beim Einsetzen von Platzhaltern kann auf die speziellen Anforderungen von
 * Dateinamen eingegangen werden. Nicht zulässige Sonderzeichen werden
 * umgewandelt und zu lange Dateinamen werden auf 255 Zeichen begrenzt, unter
 * Beibehaltung der Dateinamenserweiterung soweit diese nicht länger als 5
 * Zeichen inklusiv dem Punkt ist.<br><br>
 *
 * <b>Beispiel</b><br>
 * <pre>
 *  PlaceMarkerInserter pmi = new PlaceMarkerInserter(HashMap::new);
 *  String result = pmi.replace(
 *          "Es ist der ${sys.tag}. ${sys.monat.name} im Jahr ${sys.jahr}."
 *  );
 *  System.out.println("result = " + result);
 * </pre>
 * <p>
 * Erzeugt die Ausgabe: {@code 'Es ist der 8. Juni im Jahr 2014.'}
 * <p>
 * <!-- @date    2014-06-08 -->
 *
 * @author Zaschka, Andreas        (andreas.zaschka@mz-solutions.de)
 * @author Riebe, Moritz           (moritz.riebe@mz-solutions.de)
 */
@NotThreadSafe
public final class PlaceMarkerInserter {
    private static final String CLASS_NAME = PlaceMarkerInserter.class.getName();
    private static final Logger LOG = Logger.getLogger(CLASS_NAME);
    /**
     * Enthält vordefinierte Platzhalter.
     */
    private final Map<String, String> globalPlaceMarker;
    /**
     * Trennzeichen die den Beginn eines Platzhalters markieren.
     */
    private final String identifierStartDelimiter;
    /**
     * Trennzeichen die das Ende eines Platzhalters markieren.
     */
    private final String identifierEndDelimiter;
    /**
     * Callback wird für jeden Aufruf der {@link #replace(String)}
     * verwendet um für die aktuelle Einsetzung alle Platzhalter zu erhalten.
     */
    private Supplier<Map<String, String>> specificPlaceMarker;
    /**
     * Ersetztungs-Strategie, nach der der PlaceMarkerInserter arbeitet.
     */
    private ReplaceStrategy replaceStrategy;

    /**
     * Erzeugt einze Instanz dieser Klasse unter Nutzung selbst definierter
     * Trennzeichen für Beginn und Ende eines Platzhalters.
     * <br><br>
     * <p>
     * Mit {@code identifierStartDelimiter} und {@code identifierEndDelimiter}
     * können frei die notwendigen Zeichen(ketten) definiert werden die den
     * Beginn und das Ende eines Platzhalters markieren.<br>
     * Die Trennzeichen die ein Ende eines Platzhalters markieren, sollten
     * jedoch nicht in einem Bezeichner eines Platzhalters selbst enthalten sein.
     * <br><br>
     * <p>
     * Der Zulieferer der {@code Map<String, String>} wird für jeden
     * Einsetzungsvorgang (also pro Aufruf von
     * {@link #replace(String)}) jeweils einmalig abgerufen.
     * Die zurückgegebene Map wird zum Ersetzen der Platzhalter verwendet.
     * <br><br>
     * <p>
     * Die Ersetzungs-Strategie
     * ist per Default {@link ReplaceStrategy#REPLACE_ALL}.
     * <br><br>
     *
     * @param specificPlaceMarker      Callback-Klasse/Lambda für die Ersetzung der Platzhalter
     *                                 in einer Zeichenkette.<br>
     * @param identifierStartDelimiter Trennzeichen die den Beginn eines Platzhalters markieren.
     *                                 Länge der Zeichenkette muss größer 0 sein.<br>
     * @param identifierEndDelimiter   Trennzeichen die das Ende eines Platzhalters markieren.
     *                                 Länge der Zeichenkette muss größer 0 sein.<br>
     */
    public PlaceMarkerInserter(
            final @Nullable Supplier<Map<String, String>> specificPlaceMarker,
            final String identifierStartDelimiter,
            final String identifierEndDelimiter) {

        requireNonNull(identifierStartDelimiter, "identifierStartDelimiter");
        requireNonNull(identifierEndDelimiter, "identifierEndDelimiter");

        this.specificPlaceMarker = specificPlaceMarker;
        this.identifierStartDelimiter = identifierStartDelimiter;
        this.identifierEndDelimiter = identifierEndDelimiter;
        this.replaceStrategy = ReplaceStrategy.REPLACE_ALL;

        this.globalPlaceMarker = new HashMap<>();

        insertDefaultGlobalPlaceMarker();
    }

    /**
     * Erzeugt eine Instanz dieser Klasse mit vordefinierten Trennzeichen
     * für die Platzhalter.
     * <p>
     * Verhält sich identisch zu:
     * {@link #PlaceMarkerInserter(Supplier, String, String) }
     * <br><br>
     * <p>
     * Verwendet als vordefinierte Trennzeichen:<br>
     * <pre>
     *      identifierStartDelimiter    := ${
     *      identifierEndDelimiter      := }
     *
     *      enspr.:   ${nameDesPlatzhalters}
     * </pre>
     *
     * @param specificPlaceMarker Callback-Klasse/Lambda für die Einsetzung der Platzhalter
     *                            in einer Zeichenkatte..
     */
    public PlaceMarkerInserter(
            final Supplier<Map<String, String>> specificPlaceMarker) {

        this(specificPlaceMarker, "${", "}");
    }

    /**
     * Erzeugt eine Instanz dieser Klasse mit vordefinierten Trennzeichen
     * für die Platzhalter und keiner Ersetzungs-Map.
     * <p>
     * Identisch zu:
     * {@link #PlaceMarkerInserter(Supplier, String, String) }
     * <br><br>
     * <p>
     * Verwendet als vordefinierte Parameter:<br>
     * <pre>
     *      identifierStartDelimiter    := ${
     *      identifierEndDelimiter      := }
     *
     *      entspr.:    ${nameDesPlatzhalters}
     *
     *      specificPlaceMarker         := null
     * </pre>
     */
    public PlaceMarkerInserter() {
        this(null);
    }

    private void insertDefaultGlobalPlaceMarker() {
        final LocalDate now = LocalDate.now();

        globalPlaceMarker.put("sys.year", Integer.toString(now.getYear()));
        globalPlaceMarker.put("sys.jahr", Integer.toString(now.getYear()));

        globalPlaceMarker.put("sys.month", Integer.toString(now.getMonthValue()));
        globalPlaceMarker.put("sys.monat", Integer.toString(now.getMonthValue()));

        globalPlaceMarker.put("sys.month.name",
                now.getMonth().getDisplayName(TextStyle.FULL, Locale.US));
        globalPlaceMarker.put("sys.monat.name",
                now.getMonth().getDisplayName(TextStyle.FULL, Locale.GERMAN));

        globalPlaceMarker.put("sys.day", Integer.toString(now.getDayOfMonth()));
        globalPlaceMarker.put("sys.tag", Integer.toString(now.getDayOfMonth()));

        globalPlaceMarker.put("sys.weekday",
                now.getDayOfWeek().getDisplayName(TextStyle.FULL, Locale.US));
        globalPlaceMarker.put("sys.wochentag",
                now.getDayOfWeek().getDisplayName(TextStyle.FULL, Locale.GERMAN));

        globalPlaceMarker.put("sys.date.iso",
                now.format(DateTimeFormatter.ISO_LOCAL_DATE));
        globalPlaceMarker.put("sys.datum.iso",
                now.format(DateTimeFormatter.ISO_LOCAL_DATE));

        globalPlaceMarker.put("sys.date.din",
                now.format(DateTimeFormatter.ofPattern("dd.MM.yyyy", Locale.ENGLISH)));
        globalPlaceMarker.put("sys.datum.din",
                now.format(DateTimeFormatter.ofPattern("dd.MM.yyyy", Locale.GERMAN)));
    }

    /**
     * Setzt die Callback-Klasse zum liefern der Ersetzungsmap.
     *
     * @param specificPlaceMarkerCallback Lieferant für die Ersetzungs-Map
     */
    public void useSpecificPlaceMarker(
            final Supplier<Map<String, String>> specificPlaceMarkerCallback) {

        requireNonNull(
                specificPlaceMarkerCallback,
                "specificPlaceMarkerCallback == null");

        this.specificPlaceMarker = specificPlaceMarkerCallback;
    }

    /**
     * Setzt die Ersetzungsmap für den nächsten Ersetzungsvorang.
     *
     * @param specificPlaceMarkerMap Ersetzungs-Map
     */
    public void useSpecificPlaceMarker(
            final Map<String, String> specificPlaceMarkerMap) {

        requireNonNull(
                specificPlaceMarkerMap,
                "specificPlaceMarkerMap == null");

        this.specificPlaceMarker = () -> specificPlaceMarkerMap;
    }

    /**
     * Setzt die Ersetzungs-Strategie für die nächsten Arbeiten.
     * Solange, bis diese geändert wird.
     *
     * @param replaceStrategy Anzuwendende Ersetzungsstrategie (ist vorbelegt
     *                        mit {@link ReplaceStrategy#REPLACE_ALL}).
     */
    public void useReplaceStrategy(final ReplaceStrategy replaceStrategy) {
        requireNonNull(replaceStrategy, "replaceStrategy == null");

        this.replaceStrategy = replaceStrategy;
    }

    private String fillPlaceHolders(final String template) {
        requireNonNull(template, "template == null");

        return fillPlaceHolders(template, replaceStrategy);
    }

    private String fillPlaceHolders(
            final String template,
            final ReplaceStrategy strategy) {

        requireNonNull(template, "templateString == null");
        requireNonNull(strategy, "strategy == null");

        if (template.isEmpty()) {
            return template;
        }

        @Nullable final Map<String, String> localPlaceMarker;
        if (null != specificPlaceMarker) {
            localPlaceMarker = requireNonNull(
                    specificPlaceMarker.get(),
                    "specificPlaceMarker#get() lieferte null zurück beim "
                            + "Template: " + template);

        } else {
            localPlaceMarker = null; // keine spezielle Ersetzungsmap
        }

        final StringBuilder resultString = new StringBuilder(
                template.length() * 3);

        final CharacterIterator itr = new StringCharacterIterator(template);

        do {
            if (template.startsWith(identifierStartDelimiter, itr.getIndex())) {
                itr.setIndex(itr.getIndex() + identifierStartDelimiter.length());

                int indexOfEndDelimiter = template
                        .indexOf(identifierEndDelimiter, itr.getIndex());

                if (indexOfEndDelimiter == -1) {
                    throw new IllegalStateException("Kein Ende des "
                            + "Platzhalters gefunden: "
                            + template.substring(itr.getIndex()));
                }

                String identifier = template
                        .substring(itr.getIndex(), indexOfEndDelimiter);

                String value = getPlaceMarkerValue(
                        localPlaceMarker,
                        identifier,
                        strategy
                );

                resultString.append(value);

                itr.setIndex(
                        itr.getIndex()
                                + identifier.length()
                                + identifierEndDelimiter.length()
                                - 1
                );

            } else {
                resultString.append(itr.current());
            }

            itr.next();
        } while (itr.current() != DONE);

        return resultString.toString();
    }

    /**
     * Ersetzt Platzhalter in einer Zeichenkette.
     * <br><br>
     * <p>
     * Ersetzt alle in {@code template} auftreten Platzhalter durch die
     * vordefinierten Werte und durch die von {@code specificPlaceMarker}
     * zurückgelieferten Wert in deren {@code Map<String, String>}.
     *
     * @param template Vorlage mit Platzhaltern, darf nicht {@code null}
     *                 sein. Eine leere Zeichenkette wird direkt
     *                 zurückgegeben.<br>
     * @return Zeichenkette in der <b>alle</b> Platzhalter
     * ersetzt wurden.<br>
     * @throws IllegalStateException Falls ein Platzhalter nicht gefunden wurde oder wenn das Ende
     *                               eines Platzhalters nicht erkennbar/gefunden ist/wurde.<br>
     * @throws NullPointerException  Im Falle das {@code specificPlaceMarker#get() == null} ist.
     */
    public String replace(@Nonnull final String template) {
        return fillPlaceHolders(template);
    }

    /**
     * Ersetzt Platzhalter in einer Zeichenkette.
     * <br><br>
     * <p>
     * Ersetzt alle in {@code template} auftreten Platzhalter durch die
     * vordefinierten Werte und durch die von {@code specificPlaceMarker}
     * zurückgelieferten Wert in deren {@code Map<String, String>}.
     * <br><br>
     *
     * @param template Vorlage mit Platzhaltern, darf nicht {@code null}
     *                 sein. Eine leere Zeichenkette wird direkt
     *                 zurückgegeben.<br>
     * @param strategy Behandlungsart im Falle eines fehlenden Wertes
     *                 der einzusetzen ist.<br>
     * @return Zeichenkette in der <b>alle</b> Platzhalter
     * ersetzt wurden.
     * @throws IllegalStateException Falls ein Platzhalter nicht gefunden wurde (nur wenn
     *                               {@code strategy == ReplaceStrategy.REPLACE_ALL}) oder wenn das
     *                               Ende eines Platzhalters nicht erkennbar/gefunden ist/wurde.<br>
     * @throws NullPointerException  Im Falle das {@code specificPlaceMarker#get() == null} ist.
     */
    public String replace(
            @Nonnull final String template,
            @Nonnull final ReplaceStrategy strategy) {

        return fillPlaceHolders(template, strategy);
    }

    /**
     * Ersetzt Platzhalter in einer Zeichenkette.
     * <br><br>
     * <p>
     * Ersetzt alle in {@code template} auftreten Platzhalter durch die
     * vordefinierten Werte und durch die von {@code specificPlaceMarker}
     * zurückgelieferten Wert in deren {@code Map<String, String>}.
     * <br><br>
     *
     * <b>Sonderbehandlung von Dateinamen</b><br>
     * Beim Ersetzungsvorgang ist sichergestellt das die zurückgelieferte
     * Zeichenkette nicht länger als 255 Zeichen ist. Sollte die aus der
     * Vorlage erzeugte Zeichenkette jedoch länger sein als 255 Zeichen,
     * wird diese bis auf 255 Zeichen abgeschnitten.<br>
     * Bei Vorhandensein einer Dateinamenserweiterung, die nicht länger als
     * 5 Zeichen ist (inklusive dem Punkt), wird diese beachtet und nicht
     * abgetrennt. Die zurückgelieferte Zeichenkette ist dann trotzdem max.
     * 255 Zeichen lang.
     *
     * @param template Vorlage mit Platzhaltern, darf nicht {@code null}
     *                 sein. Eine leere Zeichenkette wird direkt
     *                 zurückgegeben.<br>
     * @return Zeichenkette in der <b>alle</b> Platzhalter
     * ersetzt wurden.<br>
     * @throws IllegalStateException Falls ein Platzhalter nicht gefunden wurde oder wenn das Ende
     *                               eines Platzhalters nicht erkennbar/gefunden ist/wurde.<br>
     * @throws NullPointerException  Im Falle das {@code specificPlaceMarker#get() == null} ist.
     */
    public String replaceForFileName(@Nonnull final String template) {
        return replaceForFileName(template, replaceStrategy);
    }

    /**
     * Ersetzt Platzhalter in einer Zeichenkette.
     * <br><br>
     * <p>
     * Ersetzt alle in {@code template} auftreten Platzhalter durch die
     * vordefinierten Werte und durch die von {@code specificPlaceMarker}
     * zurückgelieferten Wert in deren {@code Map<String, String>}.
     * <br><br>
     *
     * <b>Sonderbehandlung von Dateinamen</b><br>
     * Beim Ersetzungsvorgang ist sichergestellt das die zurückgelieferte
     * Zeichenkette nicht länger als 255 Zeichen ist. Sollte die aus der
     * Vorlage erzeugte Zeichenkette jedoch länger sein als 255 Zeichen,
     * wird diese bis auf 255 Zeichen abgeschnitten.<br>
     * Bei Vorhandensein einer Dateinamenserweiterung, die nicht länger als
     * 5 Zeichen ist (inklusive dem Punkt), wird diese beachtet und nicht
     * abgetrennt. Die zurückgelieferte Zeichenkette ist dann trotzdem max.
     * 255 Zeichen lang.
     *
     * @param template Vorlage mit Platzhaltern, darf nicht {@code null}
     *                 sein. Eine leere Zeichenkette wird direkt
     *                 zurückgegeben.<br>
     * @param strategy Behandlungsart im Falle eines fehlenden Wertes
     *                 der einzusetzen ist.<br>
     * @return Zeichenkette in der <b>alle</b> Platzhalter
     * ersetzt wurden.
     * @throws IllegalStateException Falls ein Platzhalter nicht gefunden wurde (nur wenn
     *                               {@code strategy == ReplaceStrategy.REPLACE_ALL}) oder wenn das
     *                               Ende eines Platzhalters nicht erkennbar/gefunden ist/wurde.<br>
     * @throws NullPointerException  Im Falle das {@code specificPlaceMarker#get() == null} ist.
     */
    public String replaceForFileName(
            @Nonnull final String template,
            @Nonnull final ReplaceStrategy strategy) {

        final String result = fillPlaceHolders(template, strategy);

        final String trimmed;
        if (result.length() > WindowsFileName.MAX_NAME_LENGTH) {
            trimmed = trimToFileName(result);
        } else {
            trimmed = result;
        }

        final StringBuilder buildFileNm = new StringBuilder(trimmed.length());
        for (char ch : trimmed.toCharArray()) {
            if (isLetterForbidden(ch, WindowsFileName.FORBIDDEN_CHARS)) {
                buildFileNm.append('_');
            } else {
                buildFileNm.append(ch);
            }
        }

        return buildFileNm.toString();
    }

    /**
     * Prüft einen Text auf Vorhandensein von Platzhaltern. <br><br>
     * <p>
     * Sollte der übergebene Wert {@code text == null} sein, wird dies wie
     * ein leerer String behandelt.<br>
     * <p>
     * Bei der Suche nach Platzhaltern werden die übergebenen Platzhalter
     * gesucht.<br>
     *
     * <pre>
     * "Hallo ${welt} !"  -&gt; true  , korrekt Vorhanden
     * "Hallo ${welt  !"  -&gt; false , Ende-Zeichen fehlen
     * "Hallo welt}   !"  -&gt; false , Beginn-Zeichen fehlen
     * "Hallo ${} welt!"  -&gt; true  , korrekt, wenn auch leer
     * "Hallo }Welt${ !"  -&gt; false , Falsche Reihenfolge von begin &amp; Ende
     * </pre>
     *
     * @param text Zu prüfender Text <br>
     * @return wenn {@code true}, dann sind zu ersetzende
     * Platzhalter enthalten
     */
    public boolean containsPlaceMarkers(@Nullable final String text) {
        if (null == text) {
            return false;
        }

        final int ofsBeginMarker = text.indexOf(identifierStartDelimiter);
        final int ofsEndMarker = text.indexOf(identifierEndDelimiter);

        final boolean bothExists = ofsBeginMarker != -1 && ofsEndMarker != -1;

        if (!bothExists) {
            // Existiert nur einer, dann kann es kein Platzhalter sein
            return false;
        }

        // Die Beginn-Zeichen müssen VOR den End-Zeichen des Platzhalters sein
        final boolean beginBeforeEnd = ofsBeginMarker < ofsEndMarker;

        return beginBeforeEnd;
    }

    /**
     * Durchsucht eine Vorlage nach Platzhaltern und gibt alle Platzhalter
     * ohne Trennzeichen (also ohne Beginn und Ende) als Set zurück. <br><br>
     * <p>
     * Durchsucht die Zeichenkette systematisch nach Platzhaltern. Dabei wird
     * die Zeichenkette nicht vollständig korrekt geparsed. Es werden lediglich
     * alle Zeichenketten extrahiert die zwischen den Trennzeichen stehen.
     *
     * @param template Vorlage, bei {@code null == ""} <br>
     * @return Menge aller Platzhalter; nie {@code null}
     */
    public Set<String> findAllPlaceMarker(@Nullable final String template) {
        final String pTemplate = (null == template) ? "" : template;

        if (pTemplate.isEmpty()) {
            return Collections.EMPTY_SET;
        }

        final String start = this.identifierStartDelimiter;
        final String end = this.identifierEndDelimiter;

        final int lenStart = this.identifierStartDelimiter.length();
        final int lenEnd = this.identifierEndDelimiter.length();

        final Set<String> resultSet = new HashSet<>();
        int lastIndex = 0, ofsBeginMarker, ofsEndMarker;

        while ((ofsBeginMarker = pTemplate.indexOf(start, lastIndex)) != -1) {

            ofsBeginMarker += lenStart;
            ofsEndMarker = pTemplate.indexOf(end, ofsBeginMarker);

            if (ofsEndMarker == -1) {
                break; // Kein Treffer
            }

            lastIndex = ofsEndMarker + lenEnd;

            resultSet.add(pTemplate.substring(ofsBeginMarker, ofsEndMarker));
        }

        return resultSet;
    }

    private String getPlaceMarkerValue(
            final Map<String, String> localPlaceHolder,
            final String identifier,
            final ReplaceStrategy strategy) {

        String globalValue = globalPlaceMarker.get(identifier);
        if (null != globalValue) {
            return globalValue;
        }

        if (null != localPlaceHolder) {
            String localValue = localPlaceHolder.get(identifier);
            if (null != localValue) {
                return localValue;
            }
        }

        switch (strategy) {
            case REPLACE_ALL:
                throw new IllegalStateException("Wert für Platzhalter \'"
                        + identifier + "\' nicht gefunden");

            case IGNORE_MISSING:
                return identifierStartDelimiter
                        + identifier
                        + identifierEndDelimiter;

            case REMOVE_MISSING:
                return "";

            default:
                throw new IllegalStateException(
                        "(Intern) ReplaceStrategy#" + strategy.name()
                                + " Implementierung fehlt"
                );
        }

    }

    private String trimToFileName(final String input) {
        final int MAX_LEN_FILE_NAME_EXTENSION = 5;

        final int indexOfExtension = input
                .indexOf(".", input.length() - MAX_LEN_FILE_NAME_EXTENSION);

        final boolean noFileExtension = indexOfExtension == -1;

        if (noFileExtension) {
            return input.substring(0, MAX_NAME_LENGTH);
        }

        final int sizeOfExtension = input.length() - indexOfExtension;

        return input
                .substring(0, MAX_NAME_LENGTH - sizeOfExtension)
                .concat(input.substring(indexOfExtension));
    }

    /**
     * Beschreibt die Art, wie Platzhalter in Vorlagen eingefügt werden sollen.
     */
    public static enum ReplaceStrategy {
        /**
         * Alle Platzhalter müssen ersetzt werden; wird für ein Platzhalter
         * kein zu ersetzender Wert gefunden, wird eine Exception geworfen.
         * <br><br>
         * <p>
         * Entspricht dem Standardverhalten bei Aufrufen.<br><br>
         *
         * <b>Beispiel:</b><br>
         * <pre>
         *  "${namespace.falschGesHrieben}" -&gt; führt zu Exception
         * </pre>
         */
        REPLACE_ALL,

        /**
         * Platzhalter für die kein zu ersetzender Wert gefunden wird, werden
         * entsprechend im Resultat stehen gelassen.
         * <br><br>
         *
         * <b>Beispiel:</b><br>
         * <pre>
         * "${namespace.falschGesHrieben}" -&gt; "${namespace.falschGesHrieben}"
         * </pre>
         */
        IGNORE_MISSING,

        /**
         * Platzhalter für die kein zu ersetzender Wert gefunden wurde, werden
         * aus dem Resultat entfernt.
         * <br><br>
         *
         * <b>Beispiel:</b><br>
         * <pre>
         *  "${namespace.falschGesHrieben}" -&gt; ""
         * </pre>
         */
        REMOVE_MISSING
    }

}
