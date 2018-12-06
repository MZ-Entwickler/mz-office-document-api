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

/**
 * Schnittstelle und Vorimplementierungen von Ausgabeklassen die verwendet
 * werden um erzeugte Dokumente zu speichern.
 *
 * <p>Ziel und Art der Speicherung können selbst implementiert werden, indem
 * die Schnittstelle {@link Result} implementiert wird. Alternativ können
 * vordefinierte Implementierungen genutzt werden.</p>
 * <pre>
 *  // Nutzen von fertigen Implementierungen
 *  Result resultToFile = ResultFactory.toFile(Paths.get("output.docx"));
 *
 *  // oder per Lambda-Ausdruck
 *  Result resultToLambda = (data) -&gt; System.out.println(Arrays.toString(data));
 *
 *  // oder klassisch
 *  Result resultToInnerClass = new Result() {
 *      public void writeResult(byte[] dataToWrite) throws IOException {
 *          System.out.println(Arrays.toString(dataToWrite));
 *      };
 *  };
 * </pre>
 */
package com.mz.solutions.office.result;
