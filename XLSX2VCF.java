/** Source: https://svn.apache.org/repos/asf/poi/trunk/poi-examples/src/main/java/org/apache/poi/examples/xssf/eventusermodel/XLSX2CSV.java */

package sayitobar;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.*;

@SuppressWarnings({"java:S106","java:S4823","java:S1192"})
public class XLSX2VCF {
    private final OPCPackage xlsxPackage;

    // Number of columns to read starting with leftmost
    private final int minColumns;

    // Destination for data
    private final PrintStream output;

    /**
     * Creates a new XLSX -&gt; VCF converter
     *
     * @param pkg        The XLSX package to process
     * @param output     The PrintStream to output the VCF to
     * @param minColumns The minimum number of columns to output, or -1 for no minimum
     */
    public XLSX2VCF(OPCPackage pkg, PrintStream output, int minColumns) {
        this.xlsxPackage = pkg;
        this.output = output;
        this.minColumns = minColumns;
    }
    public XLSX2VCF(OPCPackage pkg, PrintStream output) {
        this.xlsxPackage = pkg;
        this.output = output;
        this.minColumns = -1;
    }

    /**
     * Uses the XSSF Event SAX helpers to do most of the work
     *  of parsing the Sheet XML, and outputs the contents
     *  as VCF.
     */
    class SheetToVCF implements SheetContentsHandler {
        private boolean firstCellOfRow;
        private int currentRow = -1;
        private int currentCol = -1;

        private void outputMissingRows(int number) {
            for (int i=0; i < number; i++) {
                for (int j=0; j < minColumns; j++) {
                    output.append('\t');
                }
                output.append('\n');
            }
        }

        @Override
        public void startRow(int rowNum) {
            // If there were gaps, output the missing rows
            outputMissingRows(rowNum-currentRow-1);
            // Prepare for this row
            firstCellOfRow = true;
            currentRow = rowNum;
            currentCol = -1;
        }

        @Override
        public void endRow(int rowNum) {
            // Ensure the minimum number of columns
            for (int i=currentCol; i<minColumns; i++) {
                output.append('\t');
            }
            output.append('\n');
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            if (firstCellOfRow) {
                firstCellOfRow = false;
            } else {
                output.append('\t');
            }

            // gracefully handle missing CellRef here in a similar way as XSSFCell does
            if(cellReference == null) {
                cellReference = new CellAddress(currentRow, currentCol).formatAsString();
            }

            // Did we miss any cells?
            int thisCol = (new CellReference(cellReference)).getCol();
            int missedCols = thisCol - currentCol - 1;
            for (int i=0; i<missedCols; i++) {
                output.append('\t');
            }

            // no need to append anything if we do not have a value
            if (formattedValue == null) {
                return;
            }

            currentCol = thisCol;
            output.append(formattedValue);
        }

        @Override
        public void headerFooter(String s, boolean b, String s1) {

        }
    }

    /**
     * Initiates the processing of the XLS workbook file to CSV.
     *
     * @throws IOException  If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
     */
    boolean oneSheet = false;

    public void process() throws IOException, OpenXML4JException, SAXException {
        process(-1);
    }
    public void process(int sheetNum) throws IOException, OpenXML4JException, SAXException {
        oneSheet = (sheetNum >= 0);

        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader              = new XSSFReader(this.xlsxPackage);
        StylesTable styles                 = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter      = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        int index = 0;
        while (iter.hasNext()) {
            try (InputStream stream = iter.next()) {
                if (sheetNum < 0 || index == sheetNum)  // If desired sheetNum was given, read only that
                    processSheet(styles, strings, new SheetToVCF(), stream);
            }
            ++index;
        }
    }

    /**
     * Parses and shows the content of one sheet
     * using the specified styles and shared-strings tables.
     *
     * @param styles The table of styles that may be referenced by cells in the sheet
     * @param strings The table of strings that may be referenced by cells in the sheet
     * @param sheetHandler
     * @param sheetInputStream The stream to read the sheet-data from.
     */
    public void processSheet(StylesTable styles, ReadOnlySharedStringsTable strings, SheetContentsHandler sheetHandler, InputStream sheetInputStream) {
        // set emulateCSV=true on DataFormatter - it is also possible to provide a Locale
        // when POI 5.2.0 is released, you can call formatter.setUse4DigitYearsInAllDateFormats(true) to ensure all dates are formatted with 4 digit years
        DataFormatter formatter = new DataFormatter(true);
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser  = SAXHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch(ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        } catch (IOException | SAXException e) {
            e.printStackTrace();
        }
    }
}
