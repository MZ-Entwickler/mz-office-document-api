package com.mz.solutions.office.placeholders;

import com.mz.solutions.office.OfficeDocumentFactory;
import com.mz.solutions.office.instruction.DocumentProcessingInstruction;
import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.result.Result;
import com.mz.solutions.office.result.ResultFactory;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/**
 *
 * @author Moritz Riebe
 */
public class HeaderAndFooterTest extends AbstractClassPlaceholderTest {
 
    @Test
    public void testSectionsAndPages_MicrosoftOffice() {
        OfficeDocumentFactory
                .newMicrosoftOfficeInstance()
                .openDocument(selectInput("MSO_HeaderFooter_SectionsAndPages.docx"))
                .generate(
                        createDataPage(),
                        selectOutput("MSO_HeaderFooter_SectionsAndPages_Output.docx"),
                        DocumentProcessingInstruction.replaceHeaderFooterWith(createHeaderFooterPage()));
    }

    @Test
    public void testSectionsAndPages_LibreOffice() {
        OfficeDocumentFactory
                .newOpenOfficeInstance()
                .openDocument(selectInput("LOO_HeaderFooter_SectionsAndPages.odt"))
                .generate(createDataPage(), selectOutput("LOO_HeaderFooter_SectionsAndPages_Output.odt"),
                        DocumentProcessingInstruction.replaceHeaderFooterWith(createHeaderFooterPage()));
    }
    
    @Test
    public void testSimpleExistingHeaderFooter_MicrosoftOffice() {
        OfficeDocumentFactory
                .newMicrosoftOfficeInstance()
                .openDocument(selectInput("MSO_HeaderFooter_SimpleExistingHeaderFooter.docx"))
                .generate(
                        createDataPage(),
                        selectOutput("MSO_HeaderFooter_SimpleExistingHeaderFooter_Output.docx"),
                        DocumentProcessingInstruction.replaceHeaderFooterWith(createHeaderFooterPage()));
    }
    
    @Test
    public void testSimpleExistingHeaderFooter_LibreOffice() {
        OfficeDocumentFactory
                .newOpenOfficeInstance()
                .openDocument(selectInput("LOO_HeaderFooter_SimpleExistingHeaderFooter.odt"))
                .generate(createDataPage(), selectOutput("LOO_HeaderFooter_SimpleExistingHeaderFooter_Output.odt"),
                        DocumentProcessingInstruction.replaceHeaderFooterWith(createHeaderFooterPage()));
    }
    
    @Test
    public void testSimpleNoHeaderFooter_MicrosoftOffice() {
        OfficeDocumentFactory
                .newMicrosoftOfficeInstance()
                .openDocument(selectInput("MSO_HeaderFooter_SimpleNoHeaderFooter.docx"))
                .generate(createDataPage(), selectOutput("MSO_HeaderFooter_SimpleNoHeaderFooter_Output.docx"),
                        DocumentProcessingInstruction.replaceHeaderFooterWith(createHeaderFooterPage()));
    }
    
    @Test
    public void testSimpleNoHeaderFooter_LibreOffice() {
        OfficeDocumentFactory
                .newOpenOfficeInstance()
                .openDocument(selectInput("LOO_HeaderFooter_SimpleNoHeaderFooter.odt"))
                .generate(createDataPage(), selectOutput("LOO_HeaderFooter_SimpleNoHeaderFooter_Output.odt"),
                        DocumentProcessingInstruction.replaceHeaderFooterWith(createHeaderFooterPage()));
    }
    
    @Override
    protected DataPage createDataPage() {
        final DataPage page = new DataPage();
        
        page.addValue(new DataValue("ANY_VALUE", "Any-value successfully replaced."));
        
        return page;
    }
    
    protected DataPage createHeaderFooterPage() {
        final DataPage page = new DataPage();
        
        page.addValue(new DataValue("HEADER_VALUE_1", "header value 1 successfully replaced"));
        page.addValue(new DataValue("HEADER_VALUE_2", "header value 2 successfully replaced"));
        page.addValue(new DataValue("HEADER_VALUE_3", "header value 3 successfully replaced"));
        
        page.addValue(new DataValue("FOOTER_VALUE_1", "footer value 1 successfully replaced"));
        page.addValue(new DataValue("FOOTER_VALUE_2", "footer value 2 successfully replaced"));
        page.addValue(new DataValue("FOOTER_VALUE_3", "footer value 3 successfully replaced"));
        
        return page;
    }
    
    private Path selectInput(String filename) {
        return TEST_SOURCE_DIRECTORY.resolve(packageName).resolve(filename);
    }
    
    private Result selectOutput(String filename) {
        return ResultFactory.toFile(TESTS_OUTPUT_PATH.resolve(filename));
    }
    
}
