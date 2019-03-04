package com.example.demo.service;

import com.example.demo.entity.Facture;

import com.example.demo.entity.LigneFacture;
import com.itextpdf.text.Document;

import com.itextpdf.text.DocumentException;

import com.itextpdf.text.Paragraph;

import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import com.lowagie.text.pdf.PdfTable;
import org.springframework.stereotype.Service;


import java.io.IOException;

import java.io.OutputStream;


@Service

public class ExportPDFITextService {


    public void export(OutputStream os, Facture facture) throws IOException, DocumentException {
        Document document = new Document();
        PdfWriter.getInstance(document, os);
        document.open();

        /*olodocument.add(new Paragraph("Hello"));*/

        PdfPTable table = new PdfPTable(3);
        table.addCell(new PdfPCell(new Phrase("Articles")));
        table.addCell(new PdfPCell(new Phrase("Quantités")));
        table.addCell(new PdfPCell(new Phrase("Prix Unitaire")));

        for (LigneFacture ligneFacture : facture.getLigneFactures()){
            table.addCell(new PdfPCell(new Phrase(""+ligneFacture.getArticle().getLibelle())));
            table.addCell(new PdfPCell(new Phrase(""+ligneFacture.getQuantite())));
            table.addCell(new PdfPCell(new Phrase(""+ligneFacture.getArticle().getPrix())));
        }
        PdfPCell cellTotalLibelle = new PdfPCell(new Phrase("Prix Total"));
        cellTotalLibelle.setColspan(2);
        table.addCell(cellTotalLibelle);

        table.addCell(new PdfPCell((new Phrase(""+ facture.getTotal()))));

        document.add(table);

        document.close();
    }

}