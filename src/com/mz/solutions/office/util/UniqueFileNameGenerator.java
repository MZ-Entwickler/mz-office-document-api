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

import static com.mz.solutions.office.util.PlaceMarkerInserter.ReplaceStrategy.REPLACE_ALL;
import static com.mz.solutions.office.util.WindowsFileName.MAX_NAME_LENGTH;
import static com.mz.solutions.office.util.WindowsFileName.isLetterForbidden;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import static java.util.Objects.requireNonNull;
import javax.annotation.concurrent.NotThreadSafe;

/**
 * Generiert einen (Windows-kompatiblen) Dateinamen oder mehrere Dateinamen die
 * keine Doppelungen aufweisen und gemeinsam in einem Verzeichnis liegen können.
 * 
 * <p>Verhindert die doppelte Vergabe eines Dateinamens damit mehrere auto-
 * generierte Dateien in ein und demselben Verzeichnis gespeichert werden
 * können.</p>
 * 
 * <pre>
 *  Map&lt;String, String&gt; tmap = new HashMap&lt;&gt;();
 *  tmap.put("name", "Schmidt");
 *  tmap.put("datum", "2014-12-06");
 *  
 *  FileNameGenerator fnGen = new FileNameGenerator(
 *          "${datum}_Kunde_${name}.docx");
 *  
 *  for (int i = 0; i &lt; 15; i++) {
 *      System.out.println(fnGen.nextFileName(tmap));
 *  }
 * </pre>
 * 
 * <b>Ausgabe:</b><br>
 * <pre>
 * 2014_12_06_Kunde_Schmidt.docx
 * 2014_12_06_Kunde_Schmidt_2.docx
 * 2014_12_06_Kunde_Schmidt_3.docx
 * ...
 * 2014_12_06_Kunde_Schmidt_14.docx
 * 2014_12_06_Kunde_Schmidt_15.docx
 * </pre>
 * 
 * <p>Sollte der oder die generieten Dateinamen die maximal zulässige
 * Dateinamenslänge überschreiten, wird auch unter Berücksichtigung der
 * maximalen Länge, eine Nummerierung vorgenommen und ggf. die Anzahl
 * notwendiger Zeichen abgeschnitten für den Platz der Nummerierung.</p>
 * 
 * <!-- @date    2014-12-05 -->
 * @author  Riebe, Moritz       (moritz.riebe@mz-solutions.de)
 */
@NotThreadSafe
public final class UniqueFileNameGenerator {
    
    /**
     * Liste aller bereits vergebener Dateinamen (vor!) deren Umwandlung zur
     * Nummerierung. <br><br>
     * 
     * <p>Die Implementierung sieht vor das mehrfach auftretende Dateinamen auch
     * mehrfach in dieser Liste hinzugefügt (Duplikate) werden um die korrekte
     * Numerierung anhand der Anzahl zu ermitteln. Wegen String-Interning ist
     * der RAM Verbrauch dahingehend okay.</p>
     */
    private final List<String> alreadyUsedFileNames = new LinkedList<>();
    private final PlaceMarkerInserter replacer = new PlaceMarkerInserter();
    
    private final String fileNamePattern;
    
    /**
     * Erzeugt einen Dateinamens-Generator anhand einer Vorlage der Benennung
     * der Dateien.
     * 
     * <p><b>Wichtig: </b>Der Dateiname muss bereits eine
     * Dateinamenserweiterung besitzen und darf die zulässigen 255
     * Zeichenlänge überschreiten.</p>
     * 
     * @param fileNamePattern   Dateiname mit Platzhaltern der verwendet werden
     *                          soll.
     */
    public UniqueFileNameGenerator(final String fileNamePattern) {
        this.fileNamePattern = requireNonNull(fileNamePattern);
    }
    
    /**
     * Generiert einen eindeutigen und einmaligen Dateinamen anhand des im
     * Konstruktor übergebenen Dateinamens-Musters.
     * 
     * <p>Die Übergebene Ersetzungs-Map {@code replaceMap} wird zum Ersetzen der
     * Platzhalter im Dateinamens-Muster verwendet. Das Resultat nach dem
     * Ersetzungsvorgang darf die Länge samt Dateinamenserweiterung von 255
     * Zeichen überschreiten; der zurückgegebene Dateiname überschreitet diese
     * Länge jedoch nicht.</p>
     * 
     * @param replaceMap        Ersetzungs-Map zum Ausfüllen der Platzhalter
     *                          im Dateinamen
     * 
     * @return                  Einmaligener/ Eindeutiger Dateinamen; ggf. mit
     *                          einer Nummerierung am Ende
     */
    public String nextFileName(final Map<String, String> replaceMap) {
        requireNonNull(replaceMap, "replaceMap == null");
        
        replacer.useSpecificPlaceMarker(replaceMap);
        final String result = replacer.replaceForFileName(
                fileNamePattern, REPLACE_ALL).trim();
        
        final String upperResult = result.toUpperCase();
        
        // Prüfen ob dieser Name bereits verwendet wurde
        // wenn nicht, einfach zurückgeben!
        if (alreadyUsedFileNames.contains(upperResult) == false) {
            // Vor der Rückgabe EINMAL hinzufügen, weil dieser Dateiname damit
            // als 'verbraucht' registriert wird
            alreadyUsedFileNames.add(upperResult);
            return result;
        }
        
        // Auch wenn bereits vorhanden, wird der Dateiname nochmals hinzugefügt
        // (als Duplikat) um die Anzahl des Auftretens des Dateinamens
        // zählen zu können
        alreadyUsedFileNames.add(upperResult);
        
        // Zählt das wievielte mal -> bei 2 wird _2 angefügt, da der erste
        // Eintrag keine Nummerierung enthält
        final int nFile = countFileName(upperResult);

        // Länge der Zahl als Zeichenkette plus Unterstrich
        final String strNum = Integer.toString(nFile);
        final int reqLength = strNum.length() + 1;

        // Suchen des Punktes von rechts (n-1, n-2, ...)
        int indexOfDot = result.length() - 1;
        while (result.charAt(indexOfDot) != '.') {
            indexOfDot--;
        }
        
        final String nwFileName;
        final int joinLength = result.length() + reqLength;
        final boolean longerThanAllowed = joinLength >= MAX_NAME_LENGTH;
        
        if (longerThanAllowed) {
            nwFileName = result.substring(0, indexOfDot - reqLength)
                    + "_" + strNum + result.substring(indexOfDot);
            
        } else {
            nwFileName = result.substring(0, indexOfDot)
                    + "_" + strNum + result.substring(indexOfDot);
        }
        
        return nwFileName;
    }
    
    // Zählt wie häufiger der Dateiname bereits verwendet wurde unabhängig
    private int countFileName(final String fileName) {
        return (int) alreadyUsedFileNames.stream()
                .filter(name -> name.equals(fileName))
                .count();
    }
    
    /**
     * Prüft einen beliebigen Dateinamen (mit und ohne Platzhalter) ob dieser
     * eine Dateinamenserweiterung besitzt.
     * 
     * <p>Der Parameter {@code anyFileName} wird daraufhin überprüft ob er
     * (unabhängig der Groß- und Kleinschreibung) einen der Dateiendungen
     * aufweist die in {@code possibleExtensions[]} hinterlegt sind.</p>
     * 
     * @param anyFileName           zu prüfender Dateiname
     * 
     * @param possibleExtensions    Alle möglichen und zulässigen Dateinamens
     *                              endungen; Groß- und Kleinschreibung ist
     *                              unrelevant; ob mit oder ohne Punkt
     *                              beginnend ist ebenso unrelevant
     * 
     * @return                      {@code true}, wenn einer der Dateiendungen
     *                              im Dateinamen auftaucht
     */
    public static boolean hasFileNameExtension(
            final String anyFileName, final String[] possibleExtensions) {
        
        final String lowerFileName = anyFileName.trim().toLowerCase();
        
        for (String extension : possibleExtensions) {
            final String lowerExt = extension.trim().toLowerCase();
            final String cmpExtension = lowerExt.startsWith(".")
                    ? lowerExt : ('.' + lowerExt);
            
            if (lowerFileName.endsWith(cmpExtension)) {
                return true;
            }
        }
        
        return false;
    }
    
    /**
     * Filtert den übergebenen Dateinamen (egal wie lang) auf (für Windows)
     * unzulässige Zeichen und entfernt diese.
     * 
     * <p>Das Entfernen unzulässiger Sonderzeichen hat keine Längenbeschränkung
     * von 255 Zeichen, sondern führt nur das Filtern durch.</p>
     * 
     * @param fileName  Dateiname der ggf. gefiltert werden muss auf
     *                  unzulässige Zeichen.
     * 
     * @return          Dateiname ohne verbotene Sonderzeichen
     */
    public static String filterToWindowsFileName(String fileName) {
        requireNonNull(fileName, "fileName");
        
        final StringBuilder buildCleanNm = new StringBuilder(fileName.length());
        
        for (char letter : fileName.toCharArray()) {
            if (isLetterForbidden(letter, WindowsFileName.FORBIDDEN_CHARS)) {
                continue; // Nicht hinzufügen
            }
            
            buildCleanNm.append(letter); // Passt
        }
        
        return buildCleanNm.toString();
    }
    
    
}
