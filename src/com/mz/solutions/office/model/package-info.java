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

/**
 * Kurzlebiges Daten-Modell zur Strukturierung und späteren Befüllung von
 * Office Dokumenten/Reports.
 * 
 * <p><b>Grundlegendes:</b><br>
 * Das zur Befüllung von Dokumenten verwendete Daten-Modell ist ein kurzlebiges
 * Modell das nur dem Zweck der Strukturierung dient. Langfristige Datenhaltung
 * und Modellierung stehen dabei nicht im Vordergrund. Alle Modell-Klassen
 * implementieren lediglich Funktionalität zum Hinzufügen von Bezeichner-Werte-
 * Paaren, Tabellen und Tabellenzeilen. Eine Mutation dieser (Entfernen,
 * Neusortieren, ...) ist nicht vorgesehen und sollte im eigenen Domain-Modell
 * bereits erfolgt sein. Datenmodell können manuell zusammengestellt werden oder
 * aus einer externen Quelle bezogen werden. Die Unter-Packages {@code json}
 * und {@code xml} dienen als Implementierung zum Laden von vorbereiteten
 * externen Datenmodellen. Die Serialisierung (und damit auch das Laden und
 * Speichern) können alternativ ebenfalls verwendet werden.<br><br>
 * 
 * <b>Wurzelelement {@link DataPage} des Dokumenten-Objekt-Modells:</b><br>
 * Jedes Dokument (oder jede Serie von Dokumenten/Seiten) beginnen immer mit
 * einer Instanz von {@link DataPage}. Eine Instanz von {@link DataPage}
 * entspricht immer einem einzelnen Dokumentes oder einer Seite oder mehrerer
 * Seiten. Die entsprechende Interpretation ist dabei abhängig, wie das
 * Datenmodell später übergeben wird.<br><br>
 * 
 * <b>Bezeichner-Werte-Paare / Untermengen (Tabellen / Tabellenzeilen):</b><br>
 * Die unterschiedlichen Modell-Klassen übernehmen unterschiedliche Einträge
 * und Strukturen auf. Die folgende Auflistung sollte dabei verdeutlichen
 * wie die Struktur von Dokumenten ist.</p><pre>
 * 
 *  x.add(y)     | DataPage    DataTable   DataTableRow    DataValue
 *  -------------+--------------------------------------------------------------
 *  DataPage     | NEIN        JA          NEIN            JA
 *  DataTable    | NEIN        NEIN        JA              JA
 *  DataTableRow | NEIN        JA          NEIN            JA*
 *  DataValue    | NEIN        NEIN        NEIN            NEIN
 * 
 *  * Einfache DataValue's in DataTableRow, werden zum Ersetzen der Kopf und
 *    der Fußzeile verwendet; nicht zum Ersetzen von Platzhalter in den Zeilen.
 * 
 * </pre><p>Dabei ist zu beachten, dass {@link DataTable} und {@link DataValue}
 * benannte Objekte sind und entsprechend eine Bezeichnung besitzen. Jede
 * Instanz von {@link DataValue} besitzt einen Bezeichner (den Platzhalter) und
 * jede Tabelle {@link DataTable} besitzt eine unsichtbare Tabellenbezeichnung
 * die entweder direkt von Office als Name angegeben wird (so bei Open-Office)
 * oder indirekt per Textmarker in der ersten Tabellenzelle versteckt angegeben
 * wird (so bei Microsoft Office) da keine offizielle Tabellenbenennung möglich
 * ist. Zu den genauen Unterschieden im Umgang mit Platzhaltern und
 * Tabellennamen sollte die Package-Dokumentation von
 * {@code com.mz.solutions.office} herangezogen werden.<br><br>
 * 
 * <b>Beispiel:</b></p><pre>
 *  // Angaben im Hauptdokument
 *  DataPage invoiceDocument = new DataPage();
 *  invoiceDocument.addValue(new DataValue("Nachname", "Mustermann"));
 *  invoiceDocument.addValue(new DataValue("Vorname", "Max"));
 *  invoiceDocument.addValue(new DataValue("ReDatum", "2015-01-01"));
 * 
 *  // Tabelle mit den Rechnungsposten
 *  DataTable invoiceItems = new DataTable("Rechnungsposten");
 * 
 *  DataTableRow invItemRow1 = new DataTableRow();
 *  invItemRow1.addValue(new DataValue("PostenNr", "1"));
 *  invItemRow1.addValue(new DataValue("Artikel", "Pepsi Cola 1L"));
 *  invItemRow1.addValue(new DataValue("Preis", "1.70"))
 * 
 *  DataTableRow invItemRow2 = new DataTableRow();
 *  invItemRow2.addValues(   // Es gibt auch Vereinfachungen
 *          new DataValue("PostenNr", "2"),
 *          new DataValue("Artikel", "Kondensmilch"),
 *          new DataValue("Preis", "0.65"));
 * 
 *  // Hinzufügen der Zeilen; Reichenfolge ist relevant
 *  invoiceItems.addTableRow(invItemRow1);
 *  invoiceItems.addTableRow(invItemRow2);
 * 
 *  // Tabelle dem Dokument/ der Seite hinzufügen
 *  invoiceDocument.addTable(invoiceItems);
 * 
 *  // ...
 * </pre>
 */
package com.mz.solutions.office.model;
