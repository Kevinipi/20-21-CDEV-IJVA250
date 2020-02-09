package com.example.demo.controller;

import com.example.demo.entity.Article;
import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.service.impl.ArticleServiceImpl;
import com.example.demo.service.impl.ClientServiceImpl;
import com.example.demo.service.ArticleService;
import com.example.demo.service.FactureService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.servlet.HttpServletBean;
import org.springframework.web.servlet.ModelAndView;

import javax.lang.model.element.Name;
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
//Export client CSV

    @GetMapping("/clients/csv")
    public void clientsCSV (HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> clients = clientServiceImpl.findAllClients();
        writer.println("Nom" +";"+ "Prenom");
        for(int i = 0; i < clients.size(); i++) {
            Client client = clients.get(i);
            String ligne = client.getNom() +";"+ client.getPrenom();
            writer.println(ligne);
        }
    }

//Export Client XLSX

    @GetMapping("/clients/xlsx")
    public void clientsXLSX (HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/xlsx");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-clients.xlsx\"");
        PrintWriter writer = response.getWriter();
        List<Client> clients = clientServiceImpl.findAllClients();
        for(int i = 0; i < clients.size(); i++) {
            Client client = clients.get(i);
            String ligne = client.getNom() + client.getPrenom();
            writer.println(ligne);
        }
    }

//Export Articles CSV

    @GetMapping("/articles/csv")
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
}