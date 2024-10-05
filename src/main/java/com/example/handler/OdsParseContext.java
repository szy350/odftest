package com.example.handler;

import com.example.model.CurrentCell;
import com.example.model.CurrentRow;
import com.example.model.CurrentSheet;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

@Data
public class OdsParseContext {

    public OdsParseContext(Workbook workbook) {
        this.workbook = workbook;
    }

    // 对象初始化时初始化表格
    private Workbook workbook;

    private CurrentSheet sheet;

    private CurrentRow row;

    private CurrentCell cell;

    private StringBuilder valueBuilder;

    public void initSheet(Sheet sheet) {
        this.sheet = new CurrentSheet(sheet, 0);
        this.row = null;
    }

    public void initRow(Row row) {
        this.row = new CurrentRow(row, 0);
        sheet.setRowIndex(sheet.getRowIndex() + 1);
    }

    public Sheet getCurrentSheet() {
        return sheet.getSheet();
    }

    public int getCurrentSheetRowIndex() {
        return sheet.getRowIndex();
    }

    public Row getCurrentRow() {
        return row.getCurrentRow();
    }

    public int getCurrentRowCellIndex() {
        return row.getCurrentCellIndex();
    }

    public void changeRowIndex(Row row, int offset) {
        this.row.setCurrentRow(row);
        this.row.setCurrentCellIndex(0);
        sheet.setRowIndex(sheet.getRowIndex() + offset);
    }

    public void initCell(Cell cell, String cellType, String exactValue) {
        this.cell = new CurrentCell(cell, cellType, exactValue);
        this.row.setCurrentCellIndex(getCurrentRowCellIndex() + 1);
    }

    public void initValueBuilder() {
        this.valueBuilder = new StringBuilder();
    }

    public void resetValueBuilder() {
        this.valueBuilder = null;
    }
}
