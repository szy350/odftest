package com.example.handler.tabhandler;

import com.example.constant.OdsParseConstant;
import com.example.handler.OdsParseContext;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.xml.sax.Attributes;

import java.util.Optional;

public class TableTabHandler extends AbstractTabHandler {

    @Override
    public void doHandle(String uri, String localName, String qName, Attributes attributes,
        OdsParseContext odsParseContext) {
        // 表格标签页处理
        String sheetName = Optional.ofNullable(attributes.getValue(OdsParseConstant.SHEET_NAME))
            .orElse(OdsParseConstant.DEFAULT_SHEET_NAME);
        Workbook workbook = odsParseContext.getWorkbook();
        Sheet currentSheet = workbook.createSheet(sheetName);
        odsParseContext.initSheet(currentSheet);
    }
}
