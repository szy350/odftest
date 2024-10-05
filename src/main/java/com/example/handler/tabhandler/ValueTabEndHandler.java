package com.example.handler.tabhandler;

import com.example.handler.OdsParseContext;
import org.xml.sax.Attributes;

public class ValueTabEndHandler extends AbstractTabHandler {
    @Override
    public void doHandle(String uri, String localName, String qName, Attributes attributes,
        OdsParseContext odsParseContext) {
        odsParseContext.resetValueBuilder();
    }
}
