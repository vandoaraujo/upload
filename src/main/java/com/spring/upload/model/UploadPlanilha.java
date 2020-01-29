package com.spring.upload.model;

import com.fasterxml.jackson.annotation.JsonFormat;

import javax.persistence.Lob;
import java.time.LocalDate;

public class UploadPlanilha {

    @JsonFormat(shape = JsonFormat.Shape.STRING, pattern = "dd-MM-yyyy")
    private LocalDate data;

    @Lob
    private String texto;

    public String getTexto() {
        return texto;
    }

    public void setTexto(String texto) {
        this.texto = texto;
    }

    public LocalDate getData() {
        return data;
    }

    public void setData(LocalDate data) {
        this.data = data;
    }
}
