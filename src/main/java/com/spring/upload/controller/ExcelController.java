package com.spring.upload.controller;

import com.spring.upload.model.Demanda;
import com.spring.upload.model.HeaderSaida;
import com.spring.upload.model.UploadPlanilha;
import com.spring.upload.service.ExcellService;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import javax.validation.Valid;
import java.awt.*;
import java.io.IOException;
import java.util.List;

@Controller
public class ExcelController {

    @Autowired
    ExcellService excellService;

    @RequestMapping(value="/novaExtracao", method=RequestMethod.GET)
    public String getUploadForm(){
        return "novaExtracao";
    }

    @RequestMapping(value="/resultado", method= RequestMethod.GET)
    public ModelAndView getPosts(){
        ModelAndView mv = new ModelAndView("resultado");
        HeaderSaida saida = null;
        try {
            saida = excellService.extrairDados("1/16/2020");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        mv.addObject("resultado", saida);
        return mv;
    }

    @RequestMapping(value="/novaExtracao", method=RequestMethod.POST)
    public String savePost(@Valid UploadPlanilha planilha, BindingResult result, RedirectAttributes attributes){
        if(result.hasErrors()){
            attributes.addFlashAttribute("mensagem", "Verifique se os campos obrigatórios foram preenchidos!");
            return "redirect:/newupload";
        }

        return "redirect:/resultado";
    }
}
