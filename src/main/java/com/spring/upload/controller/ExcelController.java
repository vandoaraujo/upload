package com.spring.upload.controller;

import com.spring.upload.model.Demanda;
import com.spring.upload.model.HeaderSaida;
import com.spring.upload.model.UploadPlanilha;
import com.spring.upload.service.ExcellService;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import javax.validation.Valid;
import java.awt.*;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

@Controller
public class ExcelController {

    //Save the uploaded file to this folder
    private static String UPLOADED_FOLDER = "";

    @Autowired
    ExcellService excellService;

    @RequestMapping(value = "/upload", method = RequestMethod.GET)
    public String index() {
        return "upload";
    }

    @PostMapping("/upload")
    public String singleFileUpload(@RequestParam("file") MultipartFile file,
                                   RedirectAttributes redirectAttributes) {

        if (file.isEmpty()) {
            redirectAttributes.addFlashAttribute("message", "Please select a file to upload");
            return "redirect:uploadStatus";
        }

        try {

            // Get the file and save it somewhere
            byte[] bytes = file.getBytes();
            Path path = Paths.get(file.getOriginalFilename());
            Files.write(path, bytes);

            redirectAttributes.addFlashAttribute("message",
                    "Arquivo salvo com sucesso '" + file.getOriginalFilename() + "'");
            redirectAttributes.addFlashAttribute("sucesso",
                    true);

        } catch (IOException e) {
            redirectAttributes.addFlashAttribute("message",
                    "Ocorreu um erro ao efetuar o upload do arquivo " + e.getMessage());
            e.printStackTrace();
        } catch (Exception e) {
            redirectAttributes.addFlashAttribute("message",
                    "Ocorreu um erro ao efetuar o upload do arquivo " + e.getMessage());
            e.printStackTrace();
        }

        return "redirect:/uploadStatus";
    }

    @GetMapping("/uploadStatus")
    public String uploadStatus() {
        return "uploadStatus";
    }


    @RequestMapping(value = "/novaExtracao", method = RequestMethod.GET)
    public String getUploadForm() {
        return "novaExtracao";
    }

    @RequestMapping(value = "/resultado", method = RequestMethod.GET)
    public ModelAndView getPosts() throws Exception {
        ModelAndView mv = new ModelAndView("resultado");
        HeaderSaida saida = null;
        try {
            saida = excellService.extrairDados("1/16/2020");
        } catch (Exception e) {
            e.printStackTrace();
            throw new Exception(e);
        }
        mv.addObject("resultado", saida);
        return mv;
    }

    @RequestMapping(value = "/novaExtracao", method = RequestMethod.POST)
    public String savePost(@Valid UploadPlanilha planilha, BindingResult result, RedirectAttributes attributes) {
        if (result.hasErrors()) {
            attributes.addFlashAttribute("mensagem", "Verifique se os campos obrigatórios foram preenchidos!");
            return "redirect:/novaExtracao";
        }
        return "redirect:/resultado";
    }
}
