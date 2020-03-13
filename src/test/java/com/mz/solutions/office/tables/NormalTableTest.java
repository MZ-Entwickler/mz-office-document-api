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
package com.mz.solutions.office.tables;

import com.mz.solutions.office.OfficeDocument;
import com.mz.solutions.office.OfficeDocumentFactory;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataTable;
import com.mz.solutions.office.model.DataTableRow;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.result.ResultFactory;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;

public class NormalTableTest extends AbstractTableTest {

    @Test
    void testFile_NestedTables_docx() {
        final DataPage page = new DataPage();
        final DataTable tableDocument = new DataTable("T_DOCUMENT");
        final DataTable tableCustomers = new DataTable("T_CUSTOMERS");

        for (int customerIndex = 0; customerIndex < 200; customerIndex++) {
            final DataTableRow tableRow = new DataTableRow();
            tableRow.addValue(new DataValue("CUST_NO", Integer.toString((customerIndex + 1) * 10)));
            tableRow.addValue(new DataValue("CUST_NAME", randStr(19)));
            tableRow.addValue(new DataValue("CUST_ADDR", randStr(14) + "\n" + randStr(17) + "\n" + randStr(12)));

            final DataTable invoiceTable = new DataTable("T_INVOICES");
            for (int invoiceIndex = 0; invoiceIndex < (15 * Math.random()); invoiceIndex++) {
                final DataTableRow invoice = new DataTableRow();
                invoice.addValue(new DataValue("INVOICES", "2017-01-01\t12345\t" + randStr(35)));
                invoiceTable.addTableRow(invoice);
            }

            tableRow.addTable(invoiceTable);
            tableCustomers.addTableRow(tableRow);
        }

        page.addTable(tableDocument);

        final DataTableRow documentRow = new DataTableRow();
        documentRow.addTable(tableCustomers);

        tableDocument.addTableRow(documentRow);

        processWordDocument(page, NESTED_TABLES, outputPathOf(NESTED_TABLES));

    }

    @Test
    void testFile_NormalTables_docx_Repeating() {
        // Lädt eine Factory, und einmal das Dokument. Lässt es aber ZWEI mal ersetzen. Beide
        // erzeugten Dokumente müssen byte-für-byte identisch sein. Es sollte möglichst keine
        // Seiteneffekte geben.

        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(NORMAL_TABLES_DOCX);

        final ByteArrayOutputStream outFile1 = new ByteArrayOutputStream();
        final ByteArrayOutputStream outFile2 = new ByteArrayOutputStream();

        final DataPage dataPage = createDataPage();

        document.generate(dataPage, ResultFactory.toStream(outFile1));
        document.generate(dataPage, ResultFactory.toStream(outFile2));

        final byte[] dataFile1 = outFile1.toByteArray();
        final byte[] dataFile2 = outFile2.toByteArray();

        assertArrayEquals(dataFile1, dataFile2);
    }

    @Test
    void testFile_NormalTables_odt_Repeating() {
        // Lädt eine Factory, und einmal das Dokument. Lässt es aber ZWEI mal ersetzen. Beide
        // erzeugten Dokumente müssen byte-für-byte identisch sein. Es sollte möglichst keine
        // Seiteneffekte geben.

        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newOpenOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(NORMAL_TABLES_ODT);

        final ByteArrayOutputStream outFile1 = new ByteArrayOutputStream();
        final ByteArrayOutputStream outFile2 = new ByteArrayOutputStream();

        final DataPage dataPage = createDataPage();

        document.generate(dataPage, ResultFactory.toStream(outFile1));
        document.generate(dataPage, ResultFactory.toStream(outFile2));

        final byte[] dataFile1 = outFile1.toByteArray();
        final byte[] dataFile2 = outFile2.toByteArray();

        assertArrayEquals(dataFile1, dataFile2);
    }

    @Test
    void testFile_NormalTables_docx() {
        processWordDocument(createDataPage(), NORMAL_TABLES_DOCX, outputPathOf(NORMAL_TABLES_DOCX));
    }

    @Test
    void testFile_NormalTables_odt() {
        processOpenDocument(createDataPage(), NORMAL_TABLES_ODT, outputPathOf(NORMAL_TABLES_ODT));
    }


    @Test
    public void testFile_MinorBugFix_EmptyTableZeroRows_docx() {

        Path inputPath = TEST_SOURCE_DIRECTORY.resolve(packageName).resolve("MinorBugFix_EmptyTableZeroRows_MicrosoftOffice.docx");
        Path outputPath = outputPathOf(inputPath);

        processWordDocument(createDataPage(), inputPath, outputPath);
    }

    @Test
    public void testFile_MinorBugFix_EmptyTableZeroRows_odt() {

        Path inputPath = TEST_SOURCE_DIRECTORY.resolve(packageName).resolve("MinorBugFix_EmptyTableZeroRows_LibreOffice.odt");
        Path outputPath = outputPathOf(inputPath);

        processOpenDocument(createDataPage(), inputPath, outputPath);
    }

    private DataPage createDataPage() {
        final DataPage page = new DataPage();

        page.addValues(
                new DataValue("VALUE_1", "PAGE.VALUE_1"),
                new DataValue("VALUE_2", "PAGE.VALUE_2"),
                new DataValue("VALUE_3", "PAGE.VALUE_3"));

        page.addTable(createSimpleTable());
        page.addTable(createOneRowTable());
        page.addTable(createEmptyTable());
        page.addTable(createEmptyZeroRowTable());
        page.addTable(createFullTable());
        page.addTable(createEmptyTable2());

        return page;
    }

    private DataTable createSimpleTable() {
        final DataTable table = new DataTable("T_SIMPLE");

        table.addValues(
                new DataValue("VALUE_1", "PAGE.T_SIMPLE.VALUE_1"),
                new DataValue("VALUE_2", "PAGE.T_SIMPLE.VALUE_2"),
                new DataValue("VALUE_3", "PAGE.T_SIMPLE.VALUE_3"));

        table.addTableRow(new DataTableRow(
                new DataValue("VALUE_1", "PAGE.T_SIMPLE.ROW[1].VALUE_1"),
                new DataValue("VALUE_2", "PAGE.T_SIMPLE.ROW[1].VALUE_2"),
                new DataValue("VALUE_3", "PAGE.T_SIMPLE.ROW[1].VALUE_3")));

        table.addTableRow(new DataTableRow(
                new DataValue("VALUE_1", "PAGE.T_SIMPLE.ROW[2].VALUE_1"),
                new DataValue("VALUE_2", "PAGE.T_SIMPLE.ROW[2].VALUE_2"),
                new DataValue("VALUE_3", "PAGE.T_SIMPLE.ROW[2].VALUE_3")));

        table.addTableRow(new DataTableRow(
                new DataValue("VALUE_1", "PAGE.T_SIMPLE.ROW[3].VALUE_1"),
                new DataValue("VALUE_2", "PAGE.T_SIMPLE.ROW[3].VALUE_2"),
                new DataValue("VALUE_3", "PAGE.T_SIMPLE.ROW[3].VALUE_3")));

        return table;
    }

    private DataTable createOneRowTable() {
        final DataTable table = new DataTable("T_ONE_ROW");

        table.addValues(
                new DataValue("VALUE_1", "PAGE.T_ONE_ROW.VALUE_1"),
                new DataValue("VALUE_2", "PAGE.T_ONE_ROW.VALUE_2"),
                new DataValue("VALUE_3", "PAGE.T_ONE_ROW.VALUE_3"));

        table.addTableRow(new DataTableRow(
                new DataValue("VALUE_1", "PAGE.T_ONE_ROW.ROW[1].VALUE_1"),
                new DataValue("VALUE_2", "PAGE.T_ONE_ROW.ROW[1].VALUE_2"),
                new DataValue("VALUE_3", "PAGE.T_ONE_ROW.ROW[1].VALUE_3")));

        return table;
    }

    private DataTable createEmptyZeroRowTable() {
        final DataTable table = new DataTable("T_EMPTY_TABLE");
        return table;
    }

    private DataTable createEmptyTable() {
        final DataTable table = new DataTable("T_EMPTY");

        table.addValues(
                new DataValue("VALUE_1", "PAGE.T_EMPTY.VALUE_1"),
                new DataValue("VALUE_2", "PAGE.T_EMPTY.VALUE_2"),
                new DataValue("VALUE_3", "PAGE.T_EMPTY.VALUE_3"));

        return table;
    }

    private DataTable createEmptyTable2() {
        final DataTable table = new DataTable("T_EMPTY_2");

        table.addValues(
                new DataValue("VALUE_1", "PAGE.T_EMPTY_2.VALUE_1"),
                new DataValue("VALUE_2", "PAGE.T_EMPTY_2.VALUE_2"),
                new DataValue("VALUE_3", "PAGE.T_EMPTY_2.VALUE_3"));

        return table;
    }

    private DataTable createFullTable() {
        final DataTable table = new DataTable("T_FULL");

        table.addValues(
                new DataValue("VALUE_1", "PAGE.T_FULL.VALUE_1"),
                new DataValue("VALUE_2", "PAGE.T_FULL.VALUE_2"),
                new DataValue("VALUE_3", "PAGE.T_FULL.VALUE_3"));

        table.addTableRow(new DataTableRow(
                new DataValue("VALUE_1", "PAGE.T_FULL.ROW[1].VALUE_1"),
                new DataValue("VALUE_2", "PAGE.T_FULL.ROW[1].VALUE_2"),
                new DataValue("VALUE_3", "PAGE.T_FULL.ROW[1].VALUE_3")));

        table.addTableRow(new DataTableRow(
                new DataValue("VALUE_1", "PAGE.T_FULL.ROW[2].VALUE_1"),
                new DataValue("VALUE_2", "PAGE.T_FULL.ROW[2].VALUE_2"),
                new DataValue("VALUE_3", "PAGE.T_FULL.ROW[2].VALUE_3")));

        table.addTableRow(new DataTableRow(
                new DataValue("VALUE_1", "PAGE.T_FULL.ROW[3].VALUE_1"),
                new DataValue("VALUE_2", "PAGE.T_FULL.ROW[3].VALUE_2"),
                new DataValue("VALUE_3", "PAGE.T_FULL.ROW[3].VALUE_3")));

        return table;
    }

}
