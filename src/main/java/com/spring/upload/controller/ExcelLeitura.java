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
        String data = "1/16/20";
        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);
//        Map<String, RowLight> map = new HashMap();
        Iterator<Row> rowIterator = sheet.rowIterator();
        boolean primeiraLinha = true;
        List<Row> foundRows = new ArrayList<>();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            if(primeiraLinha){
                colunas = builderHeader(row);
//                colunas.forEach(s -> {
////                    System.out.println(s);
//                });
                primeiraLinha = false;
            }
            if(isRowSearched(data, row)){
                foundRows.add(row);
              //  buildRowSearched(row);
            }

            formataSaida(data, colunas, foundRows);
        }

        // Closing the workbook
        workbook.close();
    }

    /**
     * *MANUTENÇÃO E OPERAÇÃO DTR-SE*
     * *RESUMO DIÁRIO DO ATENDIMENTO - 16/01/2020*
     * EQUIPE 07:00 X 16:00
     *
     * SETD PER RELÉ Digital Modernização e testes do sistema de controle e proteção.
     * Equipe: NEANDER E GATTI
     * Viatura: 2047
     * Horário de saída: 08:30
     * Status:
     * Justificativa: DDS
     * @param colunas
     * @param foundRows
     */
    private static void formataSaida(String data, List<String> colunas, List<Row> foundRows) {
        System.out.println("MANUTENÇÃO E OPERAÇÃO DTR-SE");
        System.out.print("RESUMO DIÁRIO DO ATENDIMENTO - ");
        System.out.println(data);
        System.out.println("EQUIPE 07:00 X 16:00");

        for (int i=0 ; i< colunas.size(); i++){
            if(colunas.get(i).equals("Subestação DTR-SE")){
                System.out.println("SETD");
                System.out.print(foundRows.get(i));
            }
    //        Subestação DTR-SE	Equipamento
        }
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
