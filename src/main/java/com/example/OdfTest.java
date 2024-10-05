package com.example;

import com.alibaba.excel.util.StringUtils;
import com.example.constant.OdsParseConstant;
import com.example.handler.OdsParseContext;
import com.example.handler.OdsParseHandler;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.doc.table.OdfTableCell;
import org.odftoolkit.odfdom.doc.table.OdfTableRow;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class OdfTest {

    public static void main(String[] args) throws Exception {

        //        convertOdsToExcelByOdfTool();

        convertOdsToExcelByXmlReader();
    }

    private static void convertOdsToExcelByXmlReader() throws IOException, ParserConfigurationException, SAXException {
        Workbook workbook = new SXSSFWorkbook();
        ZipFile zipFile = new ZipFile("src/main/resources/test.ods");
        Enumeration<? extends ZipEntry> zipFileEntries = zipFile.entries();
        OdsParseHandler handler = new OdsParseHandler(new OdsParseContext(workbook));
        InputStream inputStream = null;
        InputSource inputSource = null;
        while (zipFileEntries.hasMoreElements()) {
            ZipEntry entry = zipFileEntries.nextElement();
            // 这里解析ods文件中的content.xml文件
            if (StringUtils.equals(entry.getName(), OdsParseConstant.CONTENT_XML)) {
                inputStream = zipFile.getInputStream(entry);
                inputSource = new InputSource(inputStream);
                break;
            }
        }
        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        SAXParser saxParser = saxFactory.newSAXParser();
        XMLReader xmlReader = saxParser.getXMLReader();
        xmlReader.setContentHandler(handler);
        xmlReader.parse(inputSource);

        FileOutputStream fos = new FileOutputStream("src/main/resources/test.xlsx");
        workbook.write(fos);
        inputStream.close();
        fos.flush();
    }

    private static void convertOdsToExcelByOdfTool() throws Exception {
        OdfSpreadsheetDocument odfSheet = OdfSpreadsheetDocument.loadDocument("src/main/resources/test.ods");
        List<OdfTable> odfTableList = odfSheet.getTableList();
        Workbook workbook = new XSSFWorkbook();
        for (int i = 0; i < odfTableList.size(); i++) {
            OdfTable table = odfTableList.get(i);
            table.getRowCount();

            Sheet sheet = workbook.createSheet(table.getTableName());
            for (int j = 0; j < 2000; j++) {
                OdfTableRow row = table.getRowByIndex(j);
                row.getCellCount();
                Row workBookRow = sheet.createRow(j);
                for (int k = 0; k < 20; k++) {
                    Cell workBookCell = workBookRow.createCell(k);
                    OdfTableCell cell = row.getCellByIndex(k);
                    workBookCell.setCellValue(cell.getStringValue());
                }
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(new File("src/main/resources/result.xlsx"));
        workbook.write(fileOutputStream);
        fileOutputStream.flush();
    }
}
