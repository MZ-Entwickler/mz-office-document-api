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
package com.mz.solutions.office.result;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.OpenOption;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;

/**
 * Fertige Implementierungen zum Abspeichern von Dokumenten.
 * 
 * @author  Riebe, Moritz   (moritz.riebe@mz-solutions.de)
 */
public final class ResultFactory {
    
    private ResultFactory() { }

    /**
     * Speichert das Dokument an den übergebenen Pfad.
     * 
     * @param outputFile    Zieldokument/-datei
     * 
     * @param options       Optionen die beim Schreiben beachtet werden müssen.
     * 
     * @return              {@code Result}-Implementierung
     */
    public static Result toFile(Path outputFile, OpenOption ... options) {
        return data -> Files.write(outputFile, data, options);
    }
    
    /**
     * Speichert das Dokument an den übergebenen Pfad und legt die Datei an oder
     * überschreibt eine bestehende.
     * 
     * @param outputFile    Zieldokument/-datei
     * 
     * @return              {@code Result}-Implementierung
     */
    public static Result toFile(Path outputFile) {
        return data -> Files.write(outputFile, data,
                StandardOpenOption.TRUNCATE_EXISTING,
                StandardOpenOption.CREATE);
    }
    
    /**
     * Leitet das Dokument an den übergeben Datenstrom weiter.
     * 
     * @param outputStream  Ziel-Datemstrom
     * 
     * @return              {@code Result}-Implementierung
     */
    public static Result toStream(OutputStream outputStream) {
        return data -> outputStream.write(data);
    }
    
}
