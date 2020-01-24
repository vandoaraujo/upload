package com.spring.upload.controller;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 */

public class ExcelLeitura {
	public static final String SAMPLE_XLSX_FILE_PATH = "./Programação2.xlsx";

	static List<String> colunas = null;
	static List<Map> valorContidoEmUmaLinha = new ArrayList<>();

	public static void main(String[] args) throws IOException, InvalidFormatException {

		Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
		String data = "1/16/20";
		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = sheet.rowIterator();
		boolean primeiraLinha = true;

		while (rowIterator.hasNext()) {
			Row linha = rowIterator.next();
			if (primeiraLinha) {
				colunas = builderHeader(linha);
				primeiraLinha = false;
			}
			if (isLinhaEncontrada(data, linha)) {
				// Coluna -> Valor
				Map<String, String> headerComValorCadaCelula = new HashMap<>();
				Iterator<Cell> cellIterator = linha.cellIterator();
				int indiceColuna = 0;
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					String cellValue = new DataFormatter().formatCellValue(cell);
					headerComValorCadaCelula.put(colunas.get(indiceColuna), cellValue);
					indiceColuna++;
				}
				valorContidoEmUmaLinha.add(headerComValorCadaCelula);
				// buildRowSearched(row);
			}
		}

		formataSaida(data, colunas);

		// Closing the workbook
		workbook.close();
	}

	/**
	 * *MANUTENÇÃO E OPERAÇÃO DTR-SE* *RESUMO DIÁRIO DO ATENDIMENTO - 16/01/2020*
	 * EQUIPE 07:00 X 16:00
	 *
	 * SETD PER RELÉ Digital Modernização e testes do sistema de controle e proteção.
	 * Equipe: NEANDER E GATTI
	 * Viatura: 2047
	 * Horário de saída: 08:30
	 * Status: Justificativa: DDS
	 * 
	 * @param colunas
	 * @param foundRows
	 */
	private static void formataSaida(String data, List<String> colunas) {
		System.out.println("MANUTENÇÃO E OPERAÇÃO DTR-SE");
		System.out.print("RESUMO DIÁRIO DO ATENDIMENTO - ");
		System.out.println(data);
		System.out.println("EQUIPE 07:00 X 16:00");

		for (int i = 0; i < colunas.size(); i++) {
			if (colunas.get(i).equals("Subestação DTR-SE")) {
				System.out.println("SETD");
				boolean achou = false;
				for (Map row : valorContidoEmUmaLinha) {
					Set<String> valores = row.keySet();
					for (String key : valores) {
						if(key.equals("Subestação DTR-SE")) {
							String value = (String) row.get(key);
							System.out.printf(value);
							achou = true;
							break;
						}
						
					}
					if(achou) {
						break;
					}
				}
			}
			// Subestação DTR-SE Equipamento
		}
	}

	private static List<String> builderHeader(Row row) {
		List<String> headerColunas = new ArrayList<String>();
		Iterator<Cell> cellIterator = row.cellIterator();
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			String cellValue = new DataFormatter().formatCellValue(cell);
			headerColunas.add(cellValue);
		}
		return headerColunas;
	}

	private static boolean isLinhaEncontrada(String data, Row row) {
		Iterator<Cell> cellIterator = row.cellIterator();
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			String cellValue = new DataFormatter().formatCellValue(cell);
			if (!cellValue.equals("") && (cellValue.equals(data))) {
				return true;
			}
			return false;
		}
		return false;
	}

	private static void buildRowSearched(Row row) {
		Iterator<Cell> cellIterator = row.cellIterator();
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			String cellValue = new DataFormatter().formatCellValue(cell);
			System.out.print(cellValue + "\t");
		}
		System.out.println();
	}
}
