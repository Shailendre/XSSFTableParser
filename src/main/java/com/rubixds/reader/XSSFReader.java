package com.rubixds.reader;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.List;

public class XSSFReader {

    private XSSFWorkbook workbook;
    private int sheetNumber;
    private static final SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");

    public XSSFReader(int sheetNumber) {
        this.sheetNumber = sheetNumber;
        File testWorkBook;
        try {
            testWorkBook = Paths.get(ClassLoader
                    .getSystemClassLoader()
                    .getResource("test.xlsx")
                    .toURI())
                    .toFile();
            this.workbook = new XSSFWorkbook(testWorkBook);
        } catch (URISyntaxException | InvalidFormatException | IOException e) {
            e.printStackTrace();
        }
    }

    public XSSFTable readXssfTableFromSheet(int sheetNumber) {

        XSSFSheet sheet0 = workbook.getSheetAt(sheetNumber);
        List<XSSFTable> tables = sheet0.getTables();

        /*tables.forEach(xssfTable -> {
            log.info("Table name        " + xssfTable.getDisplayName());
            log.info("Number of rows    " + xssfTable.getRowCount());
            log.info("Number of colmns  " + (xssfTable.getEndColIndex() - xssfTable.getStartColIndex() + 1) );
            log.info("Starting row index    " + xssfTable.getStartRowIndex());
            log.info("Ending row index      " + xssfTable.getEndRowIndex() + "\n");


            for (int i = xssfTable.getStartRowIndex() + 1; i <= xssfTable.getEndRowIndex(); i++) {
                XSSFRow row = sheet0.getRow(i);
                for(int j = xssfTable.getStartColIndex() ; j <= xssfTable.getEndColIndex(); j++) {
                    XSSFCell cell = row.getCell(j);
                    String stringValue = "";
                    switch (cell.getCellTypeEnum()) {
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)){
                                stringValue += dateFormat.format(cell.getDateCellValue());
                            }else{
                                stringValue += cell.getNumericCellValue();
                            }
                            break;
                        case BOOLEAN:
                            stringValue += cell.getBooleanCellValue();
                            break;
                        case FORMULA:
                            stringValue += cell.getCellFormula();
                            break;
                        default:
                            stringValue += cell.getStringCellValue();
                    }
                    log.info("Cell[" + i + "][" + j + "]    " + stringValue); }}});*/
        return tables != null && !tables.isEmpty() ? tables.get(0) : null;
    }

    public String readXSSFCellInString(XSSFCell cell) {
        String stringValue = "";
        switch (cell.getCellTypeEnum()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)){
                    stringValue += dateFormat.format(cell.getDateCellValue());
                }else{
                    stringValue += cell.getNumericCellValue();
                }
                break;
            case BOOLEAN:
                stringValue += cell.getBooleanCellValue();
                break;
            case FORMULA:
                stringValue += cell.getCellFormula();
                break;
            default:
                stringValue += cell.getStringCellValue();
        }
        return stringValue;
    }



    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public int getSheetNumber() {
        return sheetNumber;
    }

    public XSSFSheet getWorkSheet() {
        return workbook.getSheetAt(sheetNumber);
    }
}
