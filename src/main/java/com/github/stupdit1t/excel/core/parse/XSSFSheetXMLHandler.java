package com.github.stupdit1t.excel.core.parse;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.SAXException;

/**
 * 低版本没有结束sheet方法，覆盖自定义
 */
public class XSSFSheetXMLHandler extends org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler {

    private final org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler output;

    public XSSFSheetXMLHandler(StylesTable styles, CommentsTable comments, ReadOnlySharedStringsTable strings, SheetContentsHandler sheetContentsHandler, DataFormatter dataFormatter, boolean formulasNotResults) {
        super(styles, comments, strings, sheetContentsHandler, dataFormatter, formulasNotResults);
        this.output = sheetContentsHandler;
    }

    public XSSFSheetXMLHandler(StylesTable styles, ReadOnlySharedStringsTable strings, SheetContentsHandler sheetContentsHandler, DataFormatter dataFormatter, boolean formulasNotResults) {
        super(styles, strings, sheetContentsHandler, dataFormatter, formulasNotResults);
        this.output = sheetContentsHandler;
    }

    public XSSFSheetXMLHandler(StylesTable styles, ReadOnlySharedStringsTable strings, SheetContentsHandler sheetContentsHandler, boolean formulasNotResults) {
        super(styles, strings, sheetContentsHandler, formulasNotResults);
        this.output = sheetContentsHandler;
    }

    @Override
    public void endElement(String uri, String localName, String qName)
            throws SAXException {
        super.endElement(uri, localName, qName);
        if ("sheetData".equals(localName)) {
            // indicate that this sheet is now done
            ((SheetHandler) output).endSheet();
        }
    }
}
