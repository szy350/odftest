package com.example.handler.tabhandler;

import com.example.constant.OdsParseConstant;
import com.example.handler.OdsParseContext;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.xml.sax.Attributes;

public class CellTabHandler extends AbstractTabHandler {
    @Override
    public void doHandle(String uri, String localName, String qName, Attributes attributes,
        OdsParseContext odsParseContext) {

        Row row = odsParseContext.getCurrentRow();
        Cell cell = row.createCell(odsParseContext.getCurrentRowCellIndex());
        String cellType = attributes.getValue(OdsParseConstant.VALUE_TYPE);
        String exactValue = attributes.getValue(OdsParseConstant.OFFICE_VALUE);
        odsParseContext.initCell(cell, cellType, exactValue);
    }
}
