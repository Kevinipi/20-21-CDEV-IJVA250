package com.example.demo.controller;

import com.example.demo.entity.Article;
import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.service.impl.ArticleServiceImpl;
import com.example.demo.service.impl.ClientServiceImpl;
import com.example.demo.service.ArticleService;
import com.example.demo.service.FactureService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.HttpServletBean;
import org.springframework.web.servlet.ModelAndView;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Calendar;
import java.util.Date;
import javax.lang.model.element.Name;
import javax.servlet.Servlet;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.List;

/**
 * Controller principale pour affichage des clients / factures sur la page d'acceuil.
 */
@Controller
public class HomeController {

    @Autowired
    private ArticleService articleService;
    @Autowired
    private ClientServiceImpl clientServiceImpl;
    @Autowired
    private FactureService factureService;
    @Autowired
    private ArticleServiceImpl articleServiceImpl;
    private OutputStream fileOutputStream;

    public HomeController(ArticleService articleService, ClientServiceImpl clientService, FactureService factureService) {
        this.articleService = articleService;
        this.clientServiceImpl = clientService;
        this.factureService = factureService;
    }

    @GetMapping("/")
    public ModelAndView home() {
        ModelAndView modelAndView = new ModelAndView("home");

        List<Article> articles = articleService.findAll();
        modelAndView.addObject("articles", articles);

        List<Client> toto = clientServiceImpl.findAllClients();
        modelAndView.addObject("clients", toto);

        List<Facture> factures = factureService.findAllFactures();
        modelAndView.addObject("factures", factures);

        return modelAndView;
    }

//Export Articles CSV

    @RequestMapping(method = RequestMethod.GET, value = "/articles/csv")
    public void articlesCSV (HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-articles.csv\"");
        PrintWriter writer = response.getWriter();
        List<Article> articles = articleServiceImpl.findAllArticle();
        writer.println("Libellé" +";"+ "Prix");
        for(int i = 0; i < articles.size(); i++) {
            Article article = articles.get(i);
            String ligne = article.getLibelle() +";"+ article.getPrix();
            writer.println(ligne);
        }
    }

    //Export Articles XLSX

    @RequestMapping(method = RequestMethod.GET, value = "/articles/xlsx")
    public void facturesXLSX (HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/xlsx");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-articles.xlsx\"");
        PrintWriter writer = response.getWriter();
        List<Article> articles = articleServiceImpl.findAllArticle();
        writer.println("Libellé" +","+ "Prix");
        for(int i = 0; i < articles.size(); i++) {
            Article article = articles.get(i);
            String ligne = article.getLibelle() +";"+ article.getPrix();
            writer.println(ligne);
        }
    }
    //Export client CSV


    @RequestMapping(method = RequestMethod.GET, value = "/clients/csv")
    public void clientsCSV (HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> clients = clientServiceImpl.findAllClients();
        writer.println("Nom" +";"+ "Prenom" + ";" + "Age");
        for(int i = 0; i < clients.size(); i++) {
            Client client = clients.get(i);
            int age = client.getDateNaissance().until(LocalDate.now()).getYears();
            String ligne = client.getNom() +";"+ client.getPrenom() + ";" + age;
            writer.println(ligne);
        }
    }

//Export Client XLSX

@RequestMapping(method = RequestMethod.GET, value = "/clients/xlsx")
public void clientsXLSX (HttpServletRequest request, HttpServletResponse response) throws IOException {
    //ServletOutputStream os = response.getOutputStream();
    response.setContentType("application/vnd.ms-excel");
    response.setHeader("Content-Disposition", "attachment; filename=\"export-clients.xlsx\"");
    //PrintWriter writer = response.getWriter();
    List<Client> clients = clientServiceImpl.findAllClients();

    Workbook workbook = new XSSFWorkbook();
    Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);

        Cell cellPrenom = headerRow.createCell(0);
        Cell cellNom = headerRow.createCell(1);
        Cell cellAge = headerRow.createCell(2);
        cellPrenom.setCellValue("Prénom");
        cellNom.setCellValue("Nom");
        cellAge.setCellValue("Age");

        int rowNum = 1;

    for(Client client : clients) {
        Row row = sheet.createRow(rowNum ++); //A chaque ligne trouvé je recommence une ligne en dessous

        int age = client.getDateNaissance().until(LocalDate.now()).getYears();

        row.createCell(0).setCellValue(client.getPrenom());
        row.createCell(1).setCellValue(client.getNom());
        row.createCell(2).setCellValue(age);
    }
    //FileOutputStream fileOutputStream = new FileOutputStream("export-clients.xlsx");
    workbook.write(response.getOutputStream());
    //fileOutputStream.close();
    workbook.close();
}
/*
    @RequestMapping(method = RequestMethod.GET, value = "/clients/xlsx")
    public void clientsXLSX (HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/xlsx");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-clients.xlsx\"");
        PrintWriter writer = response.getWriter();
        List<Client> clients = clientServiceImpl.findAllClients();
        writer.println("Nom" +","+ "Prenom");
        for(int i = 0; i < clients.size(); i++) {
            Client client = clients.get(i);
            String ligne = client.getNom() + client.getPrenom();
            writer.println(ligne);
        }
    }

 */
}
