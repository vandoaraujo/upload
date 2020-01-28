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
	static List<Row> linhasCanceladas = new ArrayList();

	public static void main(String[] args) throws IOException, InvalidFormatException {

		Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
		String data = "1/16/20";
		String hora = "07:00 X 16:00";
		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = sheet.rowIterator();
		boolean primeiraLinha = true;

		while (rowIterator.hasNext()) {
			Row linha = rowIterator.next();
			if (primeiraLinha) {
				colunas = builderHeader(linha);
				primeiraLinha = false;
			}
			if (isDataEncontrada(data, linha) && isHoraEncontrada(hora, linha) && isRealizado(linha)) {
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
	 */
	private static void formataSaida(String data, List<String> colunas) {
		System.out.println("MANUTENÇÃO E OPERAÇÃO DTR-SE");
		System.out.print("RESUMO DIÁRIO DO ATENDIMENTO - ");
		System.out.println(data);
		System.out.println("EQUIPE 07:00 X 16:00");

		for (Map row : valorContidoEmUmaLinha) {
			System.out.print("SETD ");
			imprimeValores(colunas, "Subestação DTR-SE", row);
			imprimeValores(colunas, "Equipamento", row);
			imprimeValores(colunas, "Tipo de Equipamento", row);
			imprimeValores(colunas, "Causa/Serviço", row);
			System.out.println();
			System.out.print("Equipe: ");
			String valor = buscaColaborador(colunas, "Colaborador 1 - DTR-SE", row);
			System.out.print(valor);
			String colaborador2 = buscaColaborador(colunas, "Colaborador 2 - DTR-SE", row);
			if(colaborador2 != "") {
				System.out.print(", " + colaborador2);
			}
			String colaborador3 = buscaColaborador(colunas, "Colaborador 3 - DTR-SE", row);
			if(colaborador3 != "") {
				System.out.print(", " + colaborador3);
			}
			String colaborador4 = buscaColaborador(colunas, "Colaborador 4 - DTR-SE", row);
			if(colaborador4 != "") {
				System.out.print(", " + colaborador4);
			}
			String colaborador5 = buscaColaborador(colunas, "Observação Pós Programação", row);
			if(colaborador5 != "") {
				System.out.print(", " + colaborador5);
			}
			System.out.println();
			System.out.println("Viatura: ");
			System.out.println("Horário de Saída: ");
			System.out.println("Status: ");
			System.out.println("Justificativa: ");
			System.out.println();
		}
	}

	private static void imprimeValores(List<String> colunas, String coluna, Map row ) {
		for (int i = 0; i < colunas.size(); i++) {
			if (colunas.get(i).equals(coluna)) {
					Set<String> valores = row.keySet();
					for (String key : valores) {
						if(key.equals(coluna)) {
							String value = (String) row.get(key);
							System.out.print(value + " ");
							break;
						}
					}
			}
		}
	}

	private static String buscaColaborador(List<String> colunas, String coluna, Map row ) {
		for (int i = 0; i < colunas.size(); i++) {
			if (colunas.get(i).equals(coluna)) {
				Set<String> valores = row.keySet();
				for (String key : valores) {
					if(key.equals(coluna)) {
						return (String) row.get(key);
					}
				}
			}
		}
		return "";
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

	private static boolean isDataEncontrada(String data, Row row) {
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

	private static boolean isHoraEncontrada(String hora, Row row) {
		Iterator<Cell> cellIterator = row.cellIterator();
		boolean achou = false;
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			String cellValue = new DataFormatter().formatCellValue(cell);
			if (!cellValue.equals("") && (cellValue.equals(hora))) {
				achou = true;
				break;
			}
			achou = false;
		}
		return achou;
	}

	private static boolean isRealizado(Row row) {
		Iterator<Cell> cellIterator = row.cellIterator();
		boolean realizado = false;
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			String cellValue = new DataFormatter().formatCellValue(cell);
			if (!cellValue.equals("") && (cellValue.startsWith("Cancelado"))) {
				linhasCanceladas.add(row);
				realizado = false;
				break;
			}
			realizado = true;
		}
		return realizado;
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
