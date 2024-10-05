package com.example.handler;

import com.alibaba.excel.util.StringUtils;
import com.example.constant.OdsParseConstant;
import com.example.handler.tabhandler.*;
import com.example.model.CurrentCell;
import org.apache.poi.ss.usermodel.Cell;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.HashMap;
import java.util.Map;

public class OdsParseHandler extends DefaultHandler {

    private OdsParseContext context;

    public OdsParseHandler() {
    }

    public OdsParseHandler(OdsParseContext context) {
        this.context = context;
    }

    public Map<String, AbstractTabHandler> tabHandlerMap = new HashMap<String, AbstractTabHandler>() {
        {
            put(OdsParseConstant.TAB_TABLE, new TableTabHandler());
            put(OdsParseConstant.TAB_ROW, new RowTabHandler());
            put(OdsParseConstant.TAB_CELL, new CellTabHandler());
            put(OdsParseConstant.TAB_P, new ValueTabHandler());
        }
    };

    public Map<String, AbstractTabHandler> tabEndHandlerMap = new HashMap<String, AbstractTabHandler>() {
        {
            put(OdsParseConstant.TAB_P, new ValueTabEndHandler());
        }
    };

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        CurrentCell currentCell = context.getCell();
        if (currentCell == null) {
            return;
        }
        Cell cell = currentCell.getCurrentCell();
        String valueType = currentCell.getValueType();
        String exactValue = currentCell.getExactValue();
        if (StringUtils.isNotBlank(valueType) && StringUtils.isNotBlank(exactValue)) {
            // 单元格设置了指定的值，则用指定的值
            cell.setCellValue(exactValue);
            return;
        }
        if (context.getValueBuilder() != null) {
            context.getValueBuilder().append(ch, start, length);
            cell.setCellValue(context.getValueBuilder().toString());
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        AbstractTabHandler handler = tabEndHandlerMap.get(qName);
        if (handler == null) {
            return;
        }
        handler.doHandle(uri, localName, qName, null, context);
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        AbstractTabHandler handler = tabHandlerMap.get(qName);
        if (handler == null) {
            return;
        }
        handler.doHandle(uri, localName, qName, attributes, context);
    }
}
