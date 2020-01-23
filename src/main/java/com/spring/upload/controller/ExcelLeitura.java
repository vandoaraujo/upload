package com.spring.upload.controller;


import ch.qos.logback.core.CoreConstants;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Created by rajeevkumarsingh on 18/12/17.
 */

public class ExcelLeitura {
    public static final String SAMPLE_XLSX_FILE_PATH = "./Programação2.xlsx";

    static List<String> colunas = null;

    public static void main(String[] args) throws IOException, InvalidFormatException {

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

        // 3. Or you can use a Java 8 forEach wih lambda
        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
        workbook.forEach(sheet -> {
            System.out.println("=> " + sheet.getSheetName());
        });

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);
//        Map<String, RowLight> map = new HashMap();
        Iterator<Row> rowIterator = sheet.rowIterator();
        boolean primeiraLinha = true;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            if(primeiraLinha){
                 colunas = builderHeader(row);
                colunas.forEach(s -> {
                    System.out.println(s);
                });
                primeiraLinha = false;
            }
            if(isRowSearched("1/16/20", row)){
                buildRowSearched(row);
            }
        }

        // Closing the workbook
        workbook.close();
    }

    private static List<String> builderHeader(Row row) {
        List<String> headerColunas = new ArrayList();
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            String cellValue = new DataFormatter().formatCellValue(cell);
            headerColunas.add(cellValue);
        }
        return headerColunas;
    }

    private static boolean isRowSearched(String data, Row row ) {
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            String cellValue = new DataFormatter().formatCellValue(cell);
            if(!cellValue.equals("") && cellValue.equals(data)) {
                return true;
            }
            return false;
        }
        return false;
    }

    private static void buildRowSearched( Row row ) {
        Iterator<Cell> cellIterator = row.cellIterator();

        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            String cellValue = new DataFormatter().formatCellValue(cell);
            System.out.print(cellValue + "\t");
        }
        System.out.println();
    }
}
