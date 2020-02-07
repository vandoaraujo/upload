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
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@SuppressWarnings("ALL")
public class ExcelLeitura {
    static List<Map> valorContidoEmUmaLinha = new ArrayList<>();
    static List<Map> valorContidoEmUmaLinhaCancelados = new ArrayList<>();
    static String hora = "07:00 X 16:00";
    static String data = "1/16/20";

    public static HeaderSaida extrairDados() throws Exception {
        String SAMPLE_PERSON_DATA_FILE_PATH = "./Programacao.xlsx";
        File file = new File(SAMPLE_PERSON_DATA_FILE_PATH);
        InputStream inputStream = new FileInputStream(file);
        OPCPackage pkg = null;
        try {
            ExcelWorkSheetRowCallbackHandler sheetRowCallbackHandler =
                    new ExcelWorkSheetRowCallbackHandler(new ExcelRowContentCallback() {

                        @Override
                        public void processRow(int rowNum, Map<String, String> map) {
							if (map.get("Data").equals(data) && map.get("Status").startsWith("Cancelado")) {
							    valorContidoEmUmaLinhaCancelados.add(map);
							}
                        	else if (map.get("Data").equals(data) &&
                                    map.get("Período do Impedimento").equals(hora)) {
                                valorContidoEmUmaLinha.add(map);
//								System.out.println("rowNum=" + rowNum + ", map=" + map);
                            }
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

        } catch (RuntimeException | IOException | InvalidFormatException are) {
            are.printStackTrace();
        } finally {
            IOUtils.closeQuietly(inputStream);
            try {
                if (null != pkg) {
                    pkg.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        List<Demanda> demandas = formataSaidaDemandasRealizadas(data, hora);
        List<Demanda> demandasNaoAutorizadas = montaLayoutDemandasNaoAutorizadasCanceladas();

        HeaderSaida saida = new HeaderSaida();
        saida.setTitulo("*MANUTENÇÃO E OPERAÇÃO DTR-SE*");
        saida.setResumo("*RESUMO DIÁRIO DO ATENDIMENTO -" + data);
        saida.setHora(hora);
        saida.setAprovadas(demandas);
        saida.setCanceladas(demandasNaoAutorizadas);
        return saida;
    }

    private static List<Demanda> montaLayoutDemandasNaoAutorizadasCanceladas() {
        List<Demanda> demandasNaoAutorizadas = new ArrayList();
        /*System.out.println("SERVIÇOS NÃO AUTORIZADOS OU CANCELADOS");*/
        for (Map row : valorContidoEmUmaLinhaCancelados) {
            Demanda demandaNA = new Demanda();
            demandaNA.setTitulo("SETD ");

            Optional<String> titulo = imprimeValores("Subestação DTR-SE", row);
            demandaNA.setTitulo(demandaNA.getTitulo().concat(String.valueOf(titulo.get())).concat(" "));

            Optional<String> po = imprimeValores("PO", row);
            demandaNA.setTitulo(demandaNA.getTitulo().concat(String.valueOf(po.get())).concat(" "));

            Optional<String> causa = imprimeValores("Causa/Serviço", row);
            demandaNA.setTitulo(demandaNA.getTitulo().concat(String.valueOf(causa.get())).concat(" "));
            demandaNA.setStatus("Status: ");
            demandasNaoAutorizadas.add(demandaNA);
        }
        return demandasNaoAutorizadas;
    }

    /**
     * @param colunas
     */
    private static List<Demanda> formataSaidaDemandasRealizadas(String data, String hora) {

        List<Demanda> demandas = new ArrayList();

        for (Map row : valorContidoEmUmaLinha) {
            Demanda demanda = new Demanda();
            demanda.setTitulo("SETD ");
            Optional<String> titulo = imprimeValores("Subestação DTR-SE", row);
            demanda.setTitulo(demanda.getTitulo().concat(String.valueOf(titulo.get())).concat(" "));

            Optional<String> tituloEq = imprimeValores("Equipamento", row);
            demanda.setTitulo(demanda.getTitulo().concat(String.valueOf(tituloEq.get())).concat(" "));

            Optional<String> tipoEq = imprimeValores("Tipo de Equipamento", row);
            demanda.setTitulo(demanda.getTitulo().concat(String.valueOf(tipoEq.get())).concat(" "));

            Optional<String> po = imprimeValores("PO", row);
            demanda.setTitulo(demanda.getTitulo().concat(String.valueOf(po.get())).concat(" "));

            Optional<String> causaServico = imprimeValores("Causa/Serviço", row);
            demanda.setTitulo(demanda.getTitulo().concat(String.valueOf(causaServico.get())));

            demanda.setEquipe("Equipe: ");
            String valor = buscaColaborador("Colaborador 1 - DTR-SE", row);
            demanda.setEquipe(demanda.getEquipe().concat(valor));

            String colaborador2 = buscaColaborador("Colaborador 2 - DTR-SE", row);
            if (colaborador2 != "") {
                demanda.setEquipe(demanda.getEquipe().concat("," + colaborador2));
            }
            String colaborador3 = buscaColaborador("Colaborador 3 - DTR-SE", row);
            if (colaborador3 != "") {
                demanda.setEquipe(demanda.getEquipe().concat("," + colaborador3));
            }
            String colaborador4 = buscaColaborador("Colaborador 4 - DTR-SE", row);
            if (colaborador4 != "") {
                demanda.setEquipe(demanda.getEquipe().concat("," + colaborador4));
            }
            String colaborador5 = buscaColaborador("Observação Pós Programação", row);
            if (colaborador5 != "") {
                demanda.setEquipe(demanda.getEquipe().concat("," + colaborador5));
            }

            demanda.setViatura("Viatura: ");
            Optional<String> viatura = imprimeValores("Viatura 1 - DTR-SE", row);
            demanda.setViatura(demanda.getViatura().concat(String.valueOf(viatura.get())));
            demanda.setHorarioSaida("Horário de Saída: ");
            demanda.setStatus("Status: ");
            demanda.setJustificativa("Justificativa: ");
            demandas.add(demanda);
        }
        return demandas;
    }

    private static Optional<String> imprimeValores(String coluna, Map row) {
        Set<String> valores = row.keySet();
        for (String key : valores) {
            if (key.equals(coluna)) {
                return Optional.ofNullable((String) row.get(key));
            }
        }
        return Optional.empty();
    }

    private static String buscaColaborador(String coluna, Map row) {
        Set<String> valores = row.keySet();
        for (String key : valores) {
            if (key.equals(coluna)) {
                return (String) row.get(key);
            }
        }
        return "";
    }
}
