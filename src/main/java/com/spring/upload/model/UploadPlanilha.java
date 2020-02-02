package com.spring.upload.model;

import com.fasterxml.jackson.annotation.JsonFormat;
import java.time.LocalDate;

public class UploadPlanilha {

    private String data;

    private String planilha;

    public String getData() {
        return data;
    }

    public void setData(String data) {
        this.data = data;
    }

    public String getPlanilha() {
        return planilha;
    }

    public void setPlanilha(String planilha) {
        this.planilha = planilha;
    }
}
