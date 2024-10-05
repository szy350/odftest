package com.example.handler.tabhandler;

import com.example.handler.OdsParseContext;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.xml.sax.Attributes;

public class RowTabHandler extends AbstractTabHandler {

    @Override
    public void doHandle(String uri, String localName, String qName, Attributes attributes,
        OdsParseContext odsParseContext) {

        Sheet sheet = odsParseContext.getCurrentSheet();
        Row row = sheet.createRow(odsParseContext.getCurrentSheetRowIndex());
        if (odsParseContext.getRow() == null) {
            odsParseContext.initRow(row);
        } else {
            odsParseContext.changeRowIndex(row, 1);
        }
    }
}
