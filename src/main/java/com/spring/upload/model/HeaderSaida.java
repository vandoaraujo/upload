package com.spring.upload.model;

import java.util.List;

public class HeaderSaida {

    private String titulo;

    private String resumo;

    private String hora;

    private List<Demanda> aprovadas;

    public List<Demanda> getAprovadas() {
        return aprovadas;
    }

    public void setAprovadas(List<Demanda> aprovadas) {
        this.aprovadas = aprovadas;
    }

    public List<Demanda> getCanceladas() {
        return canceladas;
    }

    public void setCanceladas(List<Demanda> canceladas) {
        this.canceladas = canceladas;
    }

    private List<Demanda> canceladas;

    public String getTitulo() {
        return titulo;
    }

    public void setTitulo(String titulo) {
        this.titulo = titulo;
    }

    public String getResumo() {
        return resumo;
    }

    public void setResumo(String resumo) {
        this.resumo = resumo;
    }

    public String getHora() {
        return hora;
    }

    public void setHora(String hora) {
        this.hora = hora;
    }
}
