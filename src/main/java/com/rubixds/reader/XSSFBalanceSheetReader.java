package com.rubixds.reader;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;

import java.io.IOException;
import java.net.URISyntaxException;

public class XSSFBalanceSheetReader extends XSSFReader{

    public XSSFBalanceSheetReader(int sheetNumber) {
        super(sheetNumber);
    }

    class TableArea {
        int col_start, col_end, row_start, row_end;

        public int getCol_start() {
            return col_start;
        }

        public void setCol_start(int col_start) {
            this.col_start = col_start;
        }

        public int getCol_end() {
            return col_end;
        }

        public void setCol_end(int col_end) {
            this.col_end = col_end;
        }

        public int getRow_start() {
            return row_start;
        }

        public void setRow_start(int row_start) {
            this.row_start = row_start;
        }

        public int getRow_end() {
            return row_end;
        }

        public void setRow_end(int row_end) {
            this.row_end = row_end;
        }
    }

    public void extractDataFromBalanceSheet() throws InvalidFormatException, IOException, URISyntaxException {
        XSSFSheet bsSectionSheet = getWorkSheet();
        XSSFTable bsTable = bsSectionSheet.getTables().get(0);
        TableArea bsTableArea = getTableDimension(bsTable);
        XSSFRow tableHeaderRow = bsSectionSheet.getRow(bsTableArea.getRow_start());

        // print the header
        for (int i = bsTableArea.getCol_start(), j = 1; i <= bsTableArea.getCol_end(); i++, j++){
            System.out.println("Heading " + j + " " + readXSSFCellInString(tableHeaderRow.getCell(i)));
        }




    }

    private TableArea getTableDimension(XSSFTable table) {
        TableArea tableArea = new TableArea();
        tableArea.setCol_start(table.getStartColIndex());
        tableArea.setCol_end(table.getEndColIndex());
        tableArea.setRow_start(table.getStartRowIndex());
        tableArea.setRow_end(table.getEndRowIndex());
        return tableArea;
    }




}
