package com.spring.upload.controller;

import com.spring.upload.model.Demanda;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.*;

/**
 */

public class ExcelLeitura {
	public static final String SAMPLE_XLSX_FILE_PATH = "./Programação2.xlsx";

	static List<String> colunas = null;
	static List<Map> valorContidoEmUmaLinha = new ArrayList<>();
	static List<Map> valorContidoEmUmaLinhaCancelados = new ArrayList<>();
	static List<Row> linhasCanceladas = new ArrayList();

	public static void main(String[] args) throws IOException, InvalidFormatException {
		extrairDados();
	}

	public static List<Demanda> extrairDados() throws IOException, InvalidFormatException {
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
			if(isDataEncontrada(data, linha) && isCancelado(linha)){
				Map<String, String> headerComValorCadaCelula = new HashMap<>();
				Iterator<Cell> cellIterator = linha.cellIterator();
				int indiceColuna = 0;
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					String cellValue = new DataFormatter().formatCellValue(cell);
					headerComValorCadaCelula.put(colunas.get(indiceColuna), cellValue);
					indiceColuna++;
				}
				valorContidoEmUmaLinhaCancelados.add(headerComValorCadaCelula);
			}
			else if (isDataEncontrada(data, linha) && isHoraEncontrada(hora, linha)) {
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

		List<Demanda> demandas = formataSaidaDemandasRealizadas(data, hora,  colunas);
		List<Demanda> demandasNaoAutorizadas =  montaLayoutDemandasNaoAutorizadasCanceladas(colunas);
		demandas.addAll(demandasNaoAutorizadas);

		// Closing the workbook
		workbook.close();
		return demandas;
	}

	private static List<Demanda> montaLayoutDemandasNaoAutorizadasCanceladas(List<String> colunas) {
		List<Demanda> demandasNaoAutorizadas = new ArrayList();
		System.out.println("SERVIÇOS NÃO AUTORIZADOS OU CANCELADOS");
		for (Map row : valorContidoEmUmaLinhaCancelados) {
			Demanda demandaNA = new Demanda();
			demandaNA.setTitulo("SETD ");
			imprimeValores(colunas, "Subestação DTR-SE", row, demandaNA.getTitulo());
			imprimeValores(colunas, "PO", row, demandaNA.getTitulo());
			imprimeValores(colunas, "Causa/Serviço", row, demandaNA.getTitulo());
			System.out.println();
			System.out.println("Status: ");
			demandasNaoAutorizadas.add(demandaNA);
		}
		return demandasNaoAutorizadas;
	}

	/**
	 *
	 * @param colunas
	 */
	private static List<Demanda> formataSaidaDemandasRealizadas(String data, String hora, List<String> colunas) {

		List<Demanda> demandas = new ArrayList();

		System.out.println("*MANUTENÇÃO E OPERAÇÃO DTR-SE*");
		System.out.print("*RESUMO DIÁRIO DO ATENDIMENTO - ");
		System.out.println(data + "*");
		System.out.println("EQUIPE " + hora);
		System.out.println();
		System.out.println();

		for (Map row : valorContidoEmUmaLinha) {
			Demanda demanda = new Demanda();
			demanda.setTitulo("SETD ");
			imprimeValores(colunas, "Subestação DTR-SE", row, demanda.getTitulo());

			imprimeValores(colunas, "Equipamento", row, demanda.getTitulo());
			imprimeValores(colunas, "Tipo de Equipamento", row, demanda.getTitulo());
			imprimeValores(colunas, "PO", row, demanda.getTitulo());
			imprimeValores(colunas, "Causa/Serviço", row, demanda.getTitulo());
			System.out.println();
			demanda.setEquipe("Equipe: ");
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
			demanda.setViatura("Viatura: ");
			imprimeValores(colunas, "Viatura 1 - DTR-SE", row, demanda.getViatura());
			System.out.println();
			demanda.setHorarioSaida("Horário de Saída: ");
			demanda.setStatus("Status: ");
			demanda.setJustificativa("Justificativa: ");
			System.out.println();
			demandas.add(demanda);
		}
		return demandas;
	}

	private static void imprimeValores(List<String> colunas, String coluna, Map row, String value ) {
		for (int i = 0; i < colunas.size(); i++) {
			if (colunas.get(i).equals(coluna)) {
					Set<String> valores = row.keySet();
					for (String key : valores) {
						if(key.equals(coluna)) {
							value = value.concat((String) row.get(key) + " ");
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

	private static boolean isCancelado(Row row) {
		Iterator<Cell> cellIterator = row.cellIterator();
		boolean cancelado = true;
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			String cellValue = new DataFormatter().formatCellValue(cell);
			if (!cellValue.equals("") && (cellValue.startsWith("Cancelado"))) {
				cancelado = true;
				break;
			}
			cancelado = false;
		}
		return cancelado;
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
