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

import java.io.IOException;

/**
 * Schnittstelle zum Abspeichern von generierten Dokumenten.
 *
 * <p>Implementierungshinweise finden sich in der Package-Beschreibung;
 * Fertige Implementierungen sind über {@link ResultFactory} zugänglich.
 *
 * @author Riebe, Moritz   (moritz.riebe@mz-solutions.de)
 * @see ResultFactory Vorimplementierungen
 */
@FunctionalInterface
public interface Result {

    /**
     * Schreibt das übergebene Byte-Array zum gewünschten Ziel.
     *
     * <p>Im Falle einer Exception sollten alle Ressourcen freigegeben sein. Die
     * Exception kann trotzdem geworfen werden und wird hochgereicht.</p>
     *
     * @param dataToWrite Byte-Array (des Dokumentes) das zu schreiben ist
     * @throws IOException IO/Fehler; Exception darf geworfen werden
     */
    public void writeResult(byte[] dataToWrite) throws IOException;

}
