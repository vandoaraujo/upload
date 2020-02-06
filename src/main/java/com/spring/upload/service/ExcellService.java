package com.spring.upload.service;

import com.spring.upload.model.Demanda;
import com.spring.upload.model.HeaderSaida;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.List;

public interface ExcellService {

    public HeaderSaida extrairDados(String data) throws Exception;
}
