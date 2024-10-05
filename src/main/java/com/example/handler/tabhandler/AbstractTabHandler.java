package com.example.handler.tabhandler;

import com.example.handler.OdsParseContext;
import org.xml.sax.Attributes;

public abstract class AbstractTabHandler {

    public abstract void doHandle(String uri, String localName, String qName, Attributes attributes,
        OdsParseContext odsParseContext);
}
