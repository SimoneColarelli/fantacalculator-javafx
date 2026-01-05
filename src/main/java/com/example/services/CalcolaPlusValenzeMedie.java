package com.example.services;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;

import com.example.models.RoseWorkbook;
import com.example.utility.Tools;

public class CalcolaPlusValenzeMedie {
    
    private double atleticoPlusVal = 0;
    private double zarroPlusVal = 0;
    private double specialPlusVal = 0;
    private double realPlusVal = 0;
    private double spalPlusVal = 0;
    private double panzerPlusVal = 0;
    private double cammePlusVal = 0;
    private double bombePlusVal = 0;


    public double getAtleticoPlusVal() {
        return atleticoPlusVal;
    }
    public void setAtleticoPlusVal(double atleticoPlusVal) {
        this.atleticoPlusVal = atleticoPlusVal;
    }
    public double getZarroPlusVal() {
        return zarroPlusVal;
    }
    public void setZarroPlusVal(double zarroPlusVal) {
        this.zarroPlusVal = zarroPlusVal;
    }
    public double getSpecialPlusVal() {
        return specialPlusVal;
    }
    public void setSpecialPlusVal(double specialPlusVal) {
        this.specialPlusVal = specialPlusVal;
    }
    public double getRealPlusVal() {
        return realPlusVal;
    }
    public void setRealPlusVal(double realPlusVal) {
        this.realPlusVal = realPlusVal;
    }
    public double getSpalPlusVal() {
        return spalPlusVal;
    }
    public void setSpalPlusVal(double spalPlusVal) {
        this.spalPlusVal = spalPlusVal;
    }
    public double getPanzerPlusVal() {
        return panzerPlusVal;
    }
    public void setPanzerPlusVal(double panzerPlusVal) {
        this.panzerPlusVal = panzerPlusVal;
    }
    public double getCammePlusVal() {
        return cammePlusVal;
    }
    public void setCammePlusVal(double cammePlusVal) {
        this.cammePlusVal = cammePlusVal;
    }
    public double getBombePlusVal() {
        return bombePlusVal;
    }
    public void setBombePlusVal(double bombePlusVal) {
        this.bombePlusVal = bombePlusVal;
    }

    public CalcolaPlusValenzeMedie(RoseWorkbook rose) {

        for (Sheet sheet : rose) {

            String currentSheetName = sheet.getSheetName();
            if (currentSheetName.equals("COLCHONEROS") || sheet.getSheetName().equals("I_FENOMENI_DI_PERIFERIA")) continue;

            boolean newRulesPlayers = false;

            int plusValTotale = 0;
            
            for (Row row : sheet) {

                String rowStringColor = Tools.getRowStringColor((XSSFRow) row);

                //controllo che non sono a fine rosa squadra
                if (rowStringColor.equals(RoseWorkbook.BLUE_BACK)) break;

                //controllo se sto accedendo ai giocatori del nuovo ordinamento
                if (rowStringColor.equals(RoseWorkbook.YELLOW_BACK)) {
                    newRulesPlayers = true;
                    continue;
                }

                if (newRulesPlayers) {

                    int spesa = (int) row.getCell(RoseWorkbook.COLUMN_HEADERS.get("SPESA")).getNumericCellValue();
                    int valoreSvincolo = (int) row.getCell(RoseWorkbook.COLUMN_HEADERS.get("VALORE DI SVINCOLO")).getNumericCellValue();
                    
                    plusValTotale += (valoreSvincolo - spesa);

                }
            }
            
            switch(currentSheetName) {
                case "ATLETICO_ABUSIVO":
                setAtleticoPlusVal(plusValTotale);
                break;

                case "ZARRO_TEAM":
                setZarroPlusVal(plusValTotale);
                break;

                case "SPECIAL_TWO":
                setSpecialPlusVal(plusValTotale);
                break;

                case "REAL_MADRINK":
                setRealPlusVal(plusValTotale);
                break;

                case "SPAL_LETTI":
                setSpalPlusVal(plusValTotale);
                break;

                case "PANZER_TEAM":
                setPanzerPlusVal(plusValTotale);
                break;

                case "I_CAMMELLONI":
                setCammePlusVal(plusValTotale);
                break;

                case "BOMBERONI":
                setBombePlusVal(plusValTotale);
                break;
            }
        }
    }
}
