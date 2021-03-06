package com.spring.upload.service.serviceImpl;

import com.spring.upload.controller.ExcelLeitura;
import com.spring.upload.model.Demanda;
import com.spring.upload.model.HeaderSaida;
import com.spring.upload.service.ExcellService;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.Date;
import java.util.List;

@Service
public class ExcellServiceImpl implements ExcellService {


    @Override
    public HeaderSaida extrairDados(String data) throws Exception {
        return ExcelLeitura.extrairDados(data);
    }
}
