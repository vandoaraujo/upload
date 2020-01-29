package com.spring.upload.service;

import com.spring.upload.model.Demanda;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.List;

public interface ExcellService {

    public List<Demanda> extrairDados(String data) throws IOException, InvalidFormatException;
}
