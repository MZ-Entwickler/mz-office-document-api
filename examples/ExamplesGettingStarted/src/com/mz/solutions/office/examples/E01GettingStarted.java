package com.mz.solutions.office.examples;

import com.mz.solutions.office.OfficeDocument;
import com.mz.solutions.office.OfficeDocumentFactory;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataTable;
import com.mz.solutions.office.model.DataTableRow;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.result.ResultFactory;
import com.mz.solutions.office.util.PlaceMarkerInserter;
import com.mz.solutions.office.util.PlaceMarkerInserter.ReplaceStrategy;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.Optional;

public final class E01GettingStarted {
    
    public static void gettingStartedWithTheAPI() {
        // Starting with the Office-Document-Factory. For each office implementation exists
        // an office factory.
        
        // If you know the office format... create the factory directly! :-)
        final OfficeDocumentFactory msOffice = OfficeDocumentFactory.newMicrosoftOfficeInstance();
        createDocument(
                msOffice.openDocument(Paths.get("Template_Invoice.docx")),
                "INV_${INV_NUMBER}_${INV_DATE}.docx");
        
        
        // Sometimes, you don't know with format the file is (maybe without an file extension).
        // It is possible to auto detect the underlying format.
        // (But be aware.. auto detection is much slower)
        final Path libreOfficeDocPath = Paths.get("Template_Invoice.odt");
        final Optional<OfficeDocumentFactory> possibleDocFactory = OfficeDocumentFactory
                .newInstanceByDocumentType(libreOfficeDocPath);
        
        if (possibleDocFactory.isPresent() == false) {
            // Ooops.. unknown file format.
            throw new IllegalStateException("File format of " + libreOfficeDocPath
                    + " is unkonwn or can not be detected.");
        }
        
        final OfficeDocumentFactory loOffice = possibleDocFactory.get();
        createDocument(
                loOffice.openDocument(Paths.get("Template_Invoice.odt")),
                "INV_${INV_NUMBER}_${INV_DATE}.odt");
    }
    
    private static void createDocument(OfficeDocument template, String fileNamePattern) {
        final DataPage values = createPageData();
        final String fileName = createFileName(fileNamePattern, values);
        
        template.generate(
                // "values" can be a single DataPage, no data at all, or a collection of as many
                // DataPage's you want. It is everything accepted that implements Iterable oder
                // Iterator.
                values,
                
                // You may implement 'Result' yourself, or you use some pre-implemented
                // Result's with ResultFactory.
                ResultFactory.toFile(Paths.get(fileName)));
    }
    
    private static String createFileName(String pattern, DataPage values) {
        final PlaceMarkerInserter inserter = new PlaceMarkerInserter();
        inserter.useSpecificPlaceMarker(values.toMap());
        
        // You do not need to use PlaceMarkerInserter! It is just a simple utility class.
        // All placer holder classes implement #toMap(), to use it directly.
        // (btw the resulting map is read-only)
        
        return inserter.replace(pattern, ReplaceStrategy.REPLACE_ALL);
    }
    
    private static DataPage createPageData() {
        // Creating your data model. Everything starts with a DataPage
        final DataPage page = new DataPage();
        
        // Page-wide, you have your "global" values. These placeholders are replaced in the whole
        // document and tables. Header and footer in documents are just ignored.
        page.addValues(
                new DataValue("COM_NAME", "Big Business Company GmbH"),
                new DataValue("COM_STREET", "First Street 17b"),
                new DataValue("COM_ZIPCODE", "DE-01169"),
                new DataValue("COM_CITY", "Dresden"),
                
                new DataValue("INV_NUMBER", "I17B0037"),
                new DataValue("INV_DATE", LocalDate.now().toString()),
                new DataValue("INV_DUE_DATE", LocalDate.now().plusDays(14).toString()),
                new DataValue("INV_TEXT1","We do expect payment within 21 days, so please "
                        + "process this invoice within that time. There will be a 1.5% "
                        + "interest charge per month on late invoices."),
                new DataValue("INV_SUB_TOTAL", "37.00 €"),
                new DataValue("INV_SALES_TAX", "7.03 €"),
                new DataValue("INV_TOTAL", "44.03 €"),
                
                new DataValue("CUST_ADDRESS", "John Doe\nDown-Town-Street 123\nDE-01169 Dresden")
        );
        
        // A DataPage can containe as many tables as you want. Each table has its own name, within
        // its scope it must be uniqe.
        // In Libre Office or Open Office, the table name (in table properties) is directly used
        // as the matching table name.
        // In Microsoft Office ... yeah... insert a bookmark in the first cell of your table with
        // the name of the table. Sorry, we didn't found any other way to implement it.
        // btw: Each table row, can also contain nested tables!!
        final DataTable table = new DataTable("TINVOICEITEMS");
        table.addTableRow(new DataTableRow(
                new DataValue("INV_ITEM_QUANTITY", "4"),
                new DataValue("INV_ITEM_ARTICLE", "USB Stick"),
                new DataValue("INV_ITEM_UNIT_PRICE", "14.00 €"),
                new DataValue("INV_ITEM_TOTAL", "56.00 €")
        ));
        table.addTableRow(new DataTableRow(
                new DataValue("INV_ITEM_QUANTITY", "4"),
                new DataValue("INV_ITEM_ARTICLE", "Smartphone"),
                new DataValue("INV_ITEM_UNIT_PRICE", "14.00 €"),
                new DataValue("INV_ITEM_TOTAL", "56.00 €")
        ));
        table.addTableRow(new DataTableRow(
                new DataValue("INV_ITEM_QUANTITY", "4"),
                new DataValue("INV_ITEM_ARTICLE", "Smartphone"),
                new DataValue("INV_ITEM_UNIT_PRICE", "14.00 €"),
                new DataValue("INV_ITEM_TOTAL", "56.00 €")
        ));
        
        // don't forget to add the table :-)
        page.addTable(table);
        
        return page;
    }
    
}
