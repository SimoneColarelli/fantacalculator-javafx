package com.example;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.openxml4j.util.ZipSecureFile;

import com.example.models.QuotationWorkbook;
import com.example.models.RoseWorkbook;
import com.example.services.CalcolaMediaDq;
import com.example.services.CalcolaMediaPatrimoni;
import com.example.services.CalcolaPlusValenzeMedie;
import com.example.services.CalcolaTabellaValoriMercato;
import com.example.services.GenerateCsvLeghe;
import com.example.services.GenerateFileAsta;
import com.example.services.UpdateService;
import com.example.utility.Tools;

import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

public class AppController {

    public int avaragePatrimoni;

    public String dataDir = "data";

    private File fileRoseCaricato;

    private File fileQuotazioniCaricato;

    private RoseWorkbook excelRose;

    private QuotationWorkbook excelQuot;

    @FXML
    private Label meanListoneResultLabel;

    @FXML
    private Label meanRoseResultLabel;

    @FXML
    private Label fileStatusLabelRose;

    @FXML
    private Label fileStatusLabelQuotazioni;

    @FXML
    private Label caricaBilanciLabel;

    @FXML
    private void caricaRoseHandler() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Seleziona un file");
        this.fileRoseCaricato = fileChooser.showOpenDialog(new Stage());

        if (this.fileRoseCaricato != null) {
            this.fileStatusLabelRose.setText("Rose caricate");
        }
        
        this.excelRose = Tools.fileToRoseWorkbook(this.fileRoseCaricato);
    }

    @FXML
    private void caricaQuotazioniHandler() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Seleziona un file");
        this.fileQuotazioniCaricato = fileChooser.showOpenDialog(new Stage());

        if (this.fileQuotazioniCaricato != null) {
            fileStatusLabelQuotazioni.setText("Nuove Quotazioni caricate");
        }

        this.excelQuot = Tools.fileToQuotationWorkbook(this.fileQuotazioniCaricato);
    }

    @FXML 
    private void caricaBilanciHandler() throws IOException {
        
        ZipSecureFile.setMinInflateRatio(0.001); 
        
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Seleziona un file");
        File fileBilanci = fileChooser.showOpenDialog(new Stage());

        CalcolaMediaPatrimoni cmp = new CalcolaMediaPatrimoni(fileBilanci, this.fileRoseCaricato);
        this.avaragePatrimoni = cmp.getMediaPatrimoni();

        if (this.avaragePatrimoni != 0) {
            caricaBilanciLabel.setText("Bilanci caricati");
            System.out.println(this.avaragePatrimoni);
        }
    }

    @FXML
    private void calcolaNuoviValori() throws IOException {

        ZipSecureFile.setMinInflateRatio(0.001);

        File directory = new File(dataDir);
        if (!directory.exists()) {
            directory.mkdir();
        }

        RoseWorkbook rose_work = Tools.fileToRoseWorkbook(this.fileRoseCaricato);
        QuotationWorkbook quot_work = Tools.fileToQuotationWorkbook(this.fileQuotazioniCaricato);
        UpdateService updateService = new UpdateService(rose_work, quot_work);

        String xlsxFilePath = dataDir + "/Rose aggiornate.xlsx";
        String txtFilePath = dataDir + "/Giocatori non presenti in serie A.txt";

        try (FileOutputStream fileOut = new FileOutputStream(xlsxFilePath)) {
            updateService.doCompleteUpdate();
            updateService.getUpdatedRose().write(fileOut);
            System.out.println("Excel file written successfully: " + xlsxFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Creazione della lista di array e scrittura nel file TXT
        createArrayListAndWriteToFile(txtFilePath, updateService.getAlertedPlayersList());
    }

    @FXML
    private void calcolaNuoveQuotazioni() throws IOException {
        ZipSecureFile.setMinInflateRatio(0.001);

        File directory = new File(dataDir);
        if (!directory.exists()) {
            directory.mkdir();
        }

        UpdateService updateService = new UpdateService((RoseWorkbook) this.excelRose, (QuotationWorkbook) this.excelQuot);

        String xlsxFilePath = dataDir + "/Rose aggiornate.xlsx";
        String txtFilePath = dataDir + "/Giocatori non presenti in serie A.txt";

        try (FileOutputStream fileOut = new FileOutputStream(xlsxFilePath)) {
            updateService.doQuotUpdate();
            updateService.getUpdatedRose().write(fileOut);
            System.out.println("Excel file written successfully: " + xlsxFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Creazione della lista di array e scrittura nel file TXT
        createArrayListAndWriteToFile(txtFilePath, updateService.getAlertedPlayersList());
    }

    @FXML
    private void calcolaMediaDQ() throws IOException {

        QuotationWorkbook quot = Tools.fileToQuotationWorkbook(this.fileQuotazioniCaricato);
        CalcolaMediaDq calcolaMedia = new CalcolaMediaDq(quot);

        this.meanListoneResultLabel.setText(String.valueOf(calcolaMedia.getMediaDQ()));
    }

    @FXML
    private void calcolaMediaDQRose() throws IOException {

        ZipSecureFile.setMinInflateRatio(0.001);

        RoseWorkbook rose = Tools.fileToRoseWorkbook(fileRoseCaricato);
        CalcolaMediaDq calcolaMedia = new CalcolaMediaDq(rose);

        this.meanRoseResultLabel.setText(String.valueOf(calcolaMedia.getMediaDQRose()));
    }

    @FXML 
    private void calcolaPlusValenze() throws IOException {

        ZipSecureFile.setMinInflateRatio(0.001);

        File directory = new File(dataDir);
        if (!directory.exists()) {
            directory.mkdir();
        }

        RoseWorkbook rose = Tools.fileToRoseWorkbook(this.fileRoseCaricato);
        CalcolaPlusValenzeMedie cpvm = new CalcolaPlusValenzeMedie(rose);

        String txtFilePath = dataDir + "/Plus Valenze Squadre.txt";

        BufferedWriter writer = new BufferedWriter(new FileWriter(txtFilePath));
        writer.write("Atletico= " + cpvm.getAtleticoPlusVal());
        writer.newLine();
        writer.write("Zarro= " + cpvm.getZarroPlusVal());
        writer.newLine();
        writer.write("Special= " + cpvm.getSpecialPlusVal());
        writer.newLine();
        writer.write("Real= " + cpvm.getRealPlusVal());
        writer.newLine();
        writer.write("Spal= " + cpvm.getSpalPlusVal());
        writer.newLine();
        writer.write("Panzer= " + cpvm.getPanzerPlusVal());
        writer.newLine();
        writer.write("Cammelloni= " + cpvm.getCammePlusVal());
        writer.newLine();
        writer.write("Bombe= " + cpvm.getBombePlusVal());
        writer.newLine();
        writer.write("Plus valenza media delle squadre= " + ((cpvm.getAtleticoPlusVal()+cpvm.getZarroPlusVal()+cpvm.getSpecialPlusVal()+cpvm.getRealPlusVal()+cpvm.getSpalPlusVal()+cpvm.getPanzerPlusVal()+
        cpvm.getCammePlusVal()+cpvm.getBombePlusVal())/8));
        writer.close();
    }

    @FXML 
    private void creaTabelleMercato() throws IOException {

        ZipSecureFile.setMinInflateRatio(0.001);

        File directory = new File(dataDir);
        if (!directory.exists()) {
            directory.mkdir();
        }

        CalcolaTabellaValoriMercato ctvm = new CalcolaTabellaValoriMercato(this.excelRose, this.excelQuot, this.avaragePatrimoni);

        String roseFilePath = dataDir + "/Rose per mercato.xlsx";
        String quotFilePath = dataDir + "/Quotazioni per mercato.xlsx";

        try (FileOutputStream fileOutRose = new FileOutputStream(roseFilePath);
                FileOutputStream fileOutQuot = new FileOutputStream(quotFilePath)) {
            ctvm.getTabellaMercatoRose().write(fileOutRose);
            ctvm.getTabellaMercatoQuotazioni().write(fileOutQuot);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @FXML
    private void creaExcelAsta() throws IOException {
        ZipSecureFile.setMinInflateRatio(0.001);

        File directory = new File(dataDir);
        if (!directory.exists()) {
            directory.mkdir();
        }

        GenerateFileAsta gfa = new GenerateFileAsta(this.excelRose, this.excelQuot);

        String excelAstaFilePath = dataDir + "/excelAsta.xlsx";

        try (FileOutputStream fileOutExcelAsta = new FileOutputStream(excelAstaFilePath)) {
            gfa.getExcelAsta().write(fileOutExcelAsta);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @FXML
    private void exportToLegheFantacalcio() throws IOException {

        ZipSecureFile.setMinInflateRatio(0.001);

        File directory = new File(dataDir);
        if (!directory.exists()) {
            directory.mkdir();
        }

        GenerateCsvLeghe gcl = new GenerateCsvLeghe(this.excelRose, this.excelQuot);
        gcl.generateFile();
    }

    public void createArrayListAndWriteToFile(String txtFilePath, ArrayList<String> alertedPlayerList) throws IOException {
        // Creazione di una lista di array di esempio
        ArrayList<String> arrayList = alertedPlayerList;

        // Scrivi la lista di array in un file TXT
        BufferedWriter writer = new BufferedWriter(new FileWriter(txtFilePath));
        for (String alertedPlayer : arrayList) {
            writer.write(alertedPlayer);
            writer.newLine();
        }
        writer.close();
    }
}
