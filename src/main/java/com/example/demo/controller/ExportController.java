package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.ClientService;
import com.example.demo.service.FactureService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.List;
import java.util.Set;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private FactureService factureService;


    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        writer.println("Id;Nom;Prenom;Date de Naissance;Age");
        LocalDate now = LocalDate.now();
        for (Client client : allClients) {
            writer.println(
                    client.getId() + ";"
                            + client.getNom() + ";"
                            + client.getPrenom() + ";"
                            + client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")) + ";"
                            + (now.getYear() - client.getDateNaissance().getYear())
            );
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientsXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");

        Row headerRow = sheet.createRow(0);

        Cell cellHeaderId = headerRow.createCell(0);
        cellHeaderId.setCellValue("Id");

        Cell cellHeaderPrenom = headerRow.createCell(1);
        cellHeaderPrenom.setCellValue("Prénom");

        Cell cellHeaderNom = headerRow.createCell(2);
        cellHeaderNom.setCellValue("Nom");

        Cell cellHeaderDateNaissance = headerRow.createCell(3);
        cellHeaderDateNaissance.setCellValue("Date de naissance");

        int i = 1;
        for (Client client : allClients) {
            Row row = sheet.createRow(i);

            Cell cellId = row.createCell(0);
            cellId.setCellValue(client.getId());

            Cell cellPrenom = row.createCell(1);
            cellPrenom.setCellValue(client.getPrenom());

            Cell cellNom = row.createCell(2);
            cellNom.setCellValue(client.getNom());

            Cell cellDateNaissance = row.createCell(3);
            Date dateNaissance = Date.from(client.getDateNaissance().atStartOfDay(ZoneId.systemDefault()).toInstant());
            cellDateNaissance.setCellValue(dateNaissance);

            CellStyle cellStyleDate = workbook.createCellStyle();
            CreationHelper createHelper = workbook.getCreationHelper();
            cellStyleDate.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
            cellDateNaissance.setCellStyle(cellStyleDate);

            i++;
        }

        workbook.write(response.getOutputStream());
        workbook.close();

    }

    @GetMapping("/factures/xlsx")
    public void facturesXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");

        Workbook workbook = new XSSFWorkbook();

        List<Client> client = clientService.findAllClients();
        List<Facture> factures = factureService.findAllFacture();

        for (Client clients : client) {
            //boucle sur tous les clients
            Sheet sheetClient = workbook.createSheet(clients.getNom()+ " " + clients.getPrenom());
            for (Facture facture : factures) {
                //boucle sur toutes les factures pour recup toutes
                if (facture.getClient().getId() == clients.getId()) {
                    Sheet sheet = workbook.createSheet("Facture " + facture.getId()/* + " " + clients.getNom()*/);

                    Row headerRow = sheet.createRow(0);

                    //affiche le nom dans le cellules sélectionnées
                    Cell cellHeaderId = headerRow.createCell(0);
                    cellHeaderId.setCellValue("Article(s)");

                    Cell quantite = headerRow.createCell(1);
                    quantite.setCellValue("Qantité");

                    Cell prix = headerRow.createCell(2);
                    prix.setCellValue("Prix Unitaire");

                    Cell totalLigne = headerRow.createCell(3);
                    totalLigne.setCellValue("Total Ligne");

                    Set<LigneFacture> lignesf = facture.getLigneFactures();

                    int i = 1;
                    for (LigneFacture lignef : lignesf) {
                        Row l = sheet.createRow(i++);

                        Double PrixTotal = lignef.getArticle().getPrix() * lignef.getQuantite();

                        //affiche les résultats
                        Cell cellArticle = l.createCell(0);
                        cellArticle.setCellValue(lignef.getArticle().getLibelle());
                        Cell cellQuantite = l.createCell(1);
                        cellQuantite.setCellValue(lignef.getQuantite());
                        Cell cellPrixLigne = l.createCell(2);
                        cellPrixLigne.setCellValue(lignef.getArticle().getPrix());
                        Cell prixtotal = l.createCell(3);
                        prixtotal.setCellValue(PrixTotal);

                        Row totalTout = sheet.createRow(i++);
                        totalTout.createCell(2).setCellValue("Total :");
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(
                                totalTout.getRowNum(), totalTout.getRowNum(),
                                totalTout.getFirstCellNum(), (totalTout.getFirstCellNum() + 1));
                        sheet.addMergedRegion(cellRangeAddress);

                        //Style pour les cellules
                        CellStyle cellstyle = workbook.createCellStyle();
                        Font font = workbook.createFont();
                        //Style gras pour le texte
                        cellstyle.setFont(font);
                        font.setBold(true);
                        Cell cellTotal = l.createCell(4);
                        cellTotal.setCellValue(facture.getTotal());
                        cellTotal.setCellStyle(cellstyle);

                        //Style pour la couleur des cellules
                        CellStyle cellstyle1 = workbook.createCellStyle();
                        /*CellStyle rouge =*/
                    }
                }
            }
        }
        workbook.write(response.getOutputStream());
        workbook.close();
        }
    }