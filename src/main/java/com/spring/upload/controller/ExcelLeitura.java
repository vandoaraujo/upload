package com.spring.upload.controller;

import com.myjeejava.poi.ExcelReader;
import com.myjeejava.poi.ExcelRowContentCallback;
import com.myjeejava.poi.ExcelSheetCallback;
import com.myjeejava.poi.ExcelWorkSheetRowCallbackHandler;
import com.spring.upload.model.Demanda;
import com.spring.upload.model.HeaderSaida;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 */

public class ExcelLeitura {
	public static final String SAMPLE_XLSX_FILE_PATH = "./Programacao.xlsx";

	static List<String> colunas = null;
	static List<Map> valorContidoEmUmaLinha = new ArrayList<>();
	static List<Map> valorContidoEmUmaLinhaCancelados = new ArrayList<>();
	static List<Row> linhasCanceladas = new ArrayList();

	public static void main(String[] args) throws Exception {

//    String SAMPLE_PERSON_DATA_FILE_PATH = "src/test/resources/Sample-Person-Data.xlsx";

	}

	public static HeaderSaida extrairDados() throws Exception {
		String SAMPLE_PERSON_DATA_FILE_PATH = "src/test/resources/Programacao.xlsx";


		File file = new File(SAMPLE_PERSON_DATA_FILE_PATH);
		InputStream inputStream = new FileInputStream(file);

		// The package open is instantaneous, as it should be.
		OPCPackage pkg = null;
		try {
			ExcelWorkSheetRowCallbackHandler sheetRowCallbackHandler =
					new ExcelWorkSheetRowCallbackHandler(new ExcelRowContentCallback() {

						@Override
						public void processRow(int rowNum, Map<String, String> map) {

							// Do any custom row processing here, such as save
							// to database
							// Convert map values, as necessary, to dates or
							// parse as currency, etc
							System.out.println("rowNum=" + rowNum + ", map=" + map);

						}

					});

			pkg = OPCPackage.open(inputStream);

			ExcelSheetCallback sheetCallback = new ExcelSheetCallback() {
				private int sheetNumber = 0;

				@Override
				public void startSheet(int sheetNum, String sheetName) {
					this.sheetNumber = sheetNum;
					System.out.println("Started processing sheet number=" + sheetNumber
							+ " and Sheet Name is '" + sheetName + "'");
				}

				@Override
				public void endSheet() {
					System.out.println("Processing completed for sheet number=" + sheetNumber);
				}
			};


			System.out.println("Constructor: pkg, sheetRowCallbackHandler, sheetCallback");
			ExcelReader example1 = new ExcelReader(pkg, sheetRowCallbackHandler, sheetCallback);
			example1.process();

			System.out.println("\nConstructor: filePath, sheetRowCallbackHandler, sheetCallback");
			ExcelReader example2 =
					new ExcelReader(SAMPLE_PERSON_DATA_FILE_PATH, sheetRowCallbackHandler, sheetCallback);
			example2.process();

			System.out.println("\nConstructor: file, sheetRowCallbackHandler, sheetCallback");
			ExcelReader example3 = new ExcelReader(file, sheetRowCallbackHandler, null);
			example3.process();

		} catch (RuntimeException are) {
			System.out.println(are.getCause());
		} catch (InvalidFormatException ife) {
			System.out.println(ife.getCause());
		} catch (IOException ioe) {
			System.out.println(ioe.getCause());
		} finally {
			IOUtils.closeQuietly(inputStream);
			try {
				if (null != pkg) {
					pkg.close();
				}
			} catch (IOException e) {
				// just ignore IO exception
			}
		}
		return new HeaderSaida();
	}

	public static HeaderSaida extrairDados2() throws IOException, InvalidFormatException {

		HeaderSaida saida = new HeaderSaida();
//		WorkbookFactory.create()xx
		File file = new File(SAMPLE_XLSX_FILE_PATH);
		OPCPackage opcPackage = OPCPackage.open(file.getAbsolutePath());
		XSSFWorkbook workbook = new XSSFWorkbook(opcPackage);
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

		saida.setTitulo("*MANUTENÇÃO E OPERAÇÃO DTR-SE*");
		saida.setResumo("*RESUMO DIÁRIO DO ATENDIMENTO -" + data);
		saida.setHora(hora);
		saida.setAprovadas(demandas);
		saida.setCanceladas(demandasNaoAutorizadas);

//		workbook.;
		return saida;
	}

	private static List<Demanda> montaLayoutDemandasNaoAutorizadasCanceladas(List<String> colunas) {
		List<Demanda> demandasNaoAutorizadas = new ArrayList();
		/*System.out.println("SERVIÇOS NÃO AUTORIZADOS OU CANCELADOS");*/
		for (Map row : valorContidoEmUmaLinhaCancelados) {
			Demanda demandaNA = new Demanda();
			demandaNA.setTitulo("SETD ");

			Optional<String> titulo = imprimeValores(colunas, "Subestação DTR-SE", row);
			demandaNA.setTitulo(demandaNA.getTitulo().concat(String.valueOf(titulo.get())).concat(" "));

			Optional<String> po =  imprimeValores(colunas, "PO", row);
			demandaNA.setTitulo(demandaNA.getTitulo().concat(String.valueOf(po.get())).concat(" "));

			Optional<String> causa = imprimeValores(colunas, "Causa/Serviço", row);
			demandaNA.setTitulo(demandaNA.getTitulo().concat(String.valueOf(causa.get())).concat(" "));
			demandaNA.setStatus("Status: ");
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

		for (Map row : valorContidoEmUmaLinha) {
			Demanda demanda = new Demanda();
			demanda.setTitulo("SETD ");
			Optional<String> titulo = imprimeValores(colunas, "Subestação DTR-SE", row);
			demanda.setTitulo(demanda.getTitulo().concat(String.valueOf(titulo.get())).concat(" "));

			Optional<String> tituloEq = imprimeValores(colunas, "Equipamento", row);
			demanda.setTitulo(demanda.getTitulo().concat(String.valueOf(tituloEq.get())).concat(" "));

			Optional<String> tipoEq = imprimeValores(colunas, "Tipo de Equipamento", row);
			demanda.setTitulo(demanda.getTitulo().concat(String.valueOf(tipoEq.get())).concat(" "));

			Optional<String> po = imprimeValores(colunas, "PO", row);
			demanda.setTitulo(demanda.getTitulo().concat(String.valueOf(po.get())).concat(" "));

			Optional<String> causaServico = imprimeValores(colunas, "Causa/Serviço", row);
			demanda.setTitulo(demanda.getTitulo().concat(String.valueOf(causaServico.get())));

			demanda.setEquipe("Equipe: ");
			String valor = buscaColaborador(colunas, "Colaborador 1 - DTR-SE", row);
			demanda.setEquipe(demanda.getEquipe().concat(valor));

			String colaborador2 = buscaColaborador(colunas, "Colaborador 2 - DTR-SE", row);
			if(colaborador2 != "") {
				demanda.setEquipe(demanda.getEquipe().concat("," + colaborador2));
			}
			String colaborador3 = buscaColaborador(colunas, "Colaborador 3 - DTR-SE", row);
			if(colaborador3 != "") {
				demanda.setEquipe(demanda.getEquipe().concat("," + colaborador3));
			}
			String colaborador4 = buscaColaborador(colunas, "Colaborador 4 - DTR-SE", row);
			if(colaborador4 != "") {
				demanda.setEquipe(demanda.getEquipe().concat("," + colaborador4));
			}
			String colaborador5 = buscaColaborador(colunas, "Observação Pós Programação", row);
			if(colaborador5 != "") {
				demanda.setEquipe(demanda.getEquipe().concat("," + colaborador5));
			}

			demanda.setViatura("Viatura: ");
			Optional<String> viatura = imprimeValores(colunas, "Viatura 1 - DTR-SE", row);
			demanda.setViatura(demanda.getViatura().concat(String.valueOf(viatura.get())));
			demanda.setHorarioSaida("Horário de Saída: ");
			demanda.setStatus("Status: ");
			demanda.setJustificativa("Justificativa: ");
			demandas.add(demanda);
		}
		return demandas;
	}

	private static Optional<String> imprimeValores(List<String> colunas, String coluna, Map row ) {
		for (int i = 0; i < colunas.size(); i++) {
			if (colunas.get(i).equals(coluna)) {
					Set<String> valores = row.keySet();
					for (String key : valores) {
						if(key.equals(coluna)) {
							return Optional.ofNullable((String) row.get(key));
						}
					}
			}
		}
		return Optional.empty();
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
