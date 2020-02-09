package com.spring.upload.model;

import com.fasterxml.jackson.annotation.JsonFormat;
import org.springframework.format.annotation.DateTimeFormat;

import java.time.LocalDate;
import java.util.Date;
import java.util.Calendar;

public class UploadPlanilha {

    @DateTimeFormat(pattern = "yyyy-MM-dd")
    private String data;

    public String getData() {
        return data;
    }

    public void setData(String data) {
        this.data = data;
    }

    private String planilha;

    public String getPlanilha() {
        return planilha;
    }

    public void setPlanilha(String planilha) {
        this.planilha = planilha;
    }
}
