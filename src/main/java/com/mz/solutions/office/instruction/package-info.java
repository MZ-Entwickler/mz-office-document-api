/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2020,   Moritz Riebe     (moritz.riebe@mz-entwickler.de)
 *                       Andreas Zaschka  (andreas.zaschka@mz-entwickler.de)
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
 * Manueller Zugriff auf Dokumente vor der Ersetzung und nach der Ersetzung per Interceptors, sowie
 * Anweisungen um Kopf- und Fußzeilen in Dokumenten zu ersetzen..
 * 
 * <p>Dokumenten-Anweisungen können manuelle Eingriffe sein, welche vor der Ersetzung ausgeführt 
 * werden {@link DocumentInterceptorType#BEFORE_GENERATION} oder wenn gewünscht nach Ausführung der 
 * Ersetzung {@link DocumentInterceptorType#AFTER_GENERATION}.</p>
 * 
 * <p>Die Factory {@link DocumentProcessingInstruction} bietet vereinfachte statische Methoden an
 * um die häufigsten Fälle abdecken zu können.</p>
 * 
 * <pre>
 * final DocumentInterceptorFunction interceptorFunction = ...
 * final DocumentInterceptorFunction changeCustomXml = (DocumentInterceptionContext context) -&gt; {
 *     final Document xmlDocument = context.getXmlDocument();
 *     final NodeList styleNodes = xmlDocument.getElementsByTagName("custXml:customers");
 *
 *     // add/remove/change XML document
 *     // 'context' should contain all data and access you will need
 * };
 *
 * anyDocument.generate(documentData, ResultFactory.toFile(invoiceOutput),
 *         // Intercept main document part (document body)
 *         DocumentProcessingInstruction.interceptDocumentBody(
 *                 DocumentInterceptorType.BEFORE_GENERATION,  // invoke interceptor before
 *                 interceptorFunction, // change low level document function (Callback-Method)
 *                 dataMapForInterceptorFunctionHere), // data is optional
 *         // Intercept styles part of this document, maybe to change font-scaling afterwards
 *         DocumentProcessingInstruction.interceptDocumentStylesAfter(
 *                 interceptorFunction), // no data for this callback function
 *         // let us change the Custom XML Document Part (only MS Word) und fill with our data
 *         DocumentProcessingInstruction.interceptXmlDocumentPart(
 *                 "word/custom/custPropItem.xml", // our Custom XML data
 *                 DocumentInterceptorType.AFTER_GENERATION, // before or after doesn't matter
 *                 changeCustomXml));
 * </pre>
 * 
 * <p>Abgesehen von manuellen Anweisungen, können dem Ersetzungsvorgang auch erweiterte Anweisungen
 * zur Ersetzung von Kopf- und Fußzeilen mit übergeben werden. Die Ersetzung benötigt in dem Falle
 * ein eigenes Daten-Modell {@link com.mz.solutions.office.model.DataPage} da die Kopf- und
 * Fußzeilen nicht seiten-gebunden sind, sondern durch die Vorlage bestimmt sind mit dem
 * Möglichkeiten der jeweiligen Office-Variante (z.B. abweichende Fußzeilen jenachdem ob es sich um
 * eine gerade oder ungerade Seitennummer handelt).</p>
 * 
 * <pre>
 * // Header and Footer maybe exists in 'anyDocument'; if not, the library don't mind
 * final OfficeDocument anyDocument = ...
 *
 * final DataPage documentData = ...    // Values for the main document content
 * final DataPage headerData = ...      // Values to replace placeholders in the header
 * final DataPage footerData = ...      // Values to replace placeholders in the footer
 *
 * anyDocument.generate(documentData, ResultFactory.toFile(invoiceOutput),
 *         DocumentProcessingInstruction.replaceHeaderWith(headerData),
 *         DocumentProcessingInstruction.replaceFooterWith(footerData));
 * </pre>
 * 
 * Wenn Kopf- und Fußzeile aus den selben Werten ersetzt werden sollen/können, kann auch noch
 * vereinfachender eine kleinere Anweisung verwendet werden.
 * 
 * <pre>
 * final DataPage headerFooterSameData = ...    // (same) values for header AND footer
 * 
 * anyDocument.generate(documentData, ResultFactory.toFile(invoiceOutput),
 *         DocumentProcessingInstruction.replaceHeaderFooterWith(headerFooterSameData));
 * </pre>
 */
package com.mz.solutions.office.instruction;
