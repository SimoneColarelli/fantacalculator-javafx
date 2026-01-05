package com.example.services;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.example.models.QuotationWorkbook;
import com.example.models.RoseWorkbook;
import com.example.utility.Tools;

public class CalcolaMediaDq {

    private double mediaDQ = 0;

    private double mediaDQRose = 0;

    public double getMediaDQRose() {
        return mediaDQRose;
    }

    public void setMediaDQRose(double mediaDQRose) {
        this.mediaDQRose = mediaDQRose;
    }

    public double getMediaDQ() {
        return mediaDQ;
    }

    public void setMediaDQ(double mediaDQ) {
        this.mediaDQ = mediaDQ;
    }

    public CalcolaMediaDq(XSSFWorkbook quotWork) {

        double dqSum = 0;
        int counter = 0;
        XSSFSheet sheet = quotWork.getSheetAt(0);

        for (int i = QuotationWorkbook.FIRST_PLAYER_ROW; i <= sheet.getLastRowNum(); i++) {
            XSSFRow currentRow = sheet.getRow(i);

            if (currentRow.getCell(0) == null) break;

            if (currentRow.getCell(QuotationWorkbook.QUOT_COLUMN).getCellType() != CellType.NUMERIC) continue;

            int att_quot = (int) currentRow.getCell(QuotationWorkbook.QUOT_COLUMN).getNumericCellValue();
            int init_quot = (int) currentRow.getCell(QuotationWorkbook.INIT_QUOT_COLUMN).getNumericCellValue();

            if(att_quot != 1 && init_quot != 1) {
                int dq = att_quot - init_quot;
                dqSum += dq;
                counter += 1;
            }
        }

        setMediaDQ(dqSum/counter);
    }

    public CalcolaMediaDq(RoseWorkbook roseWork) {
        
        double dqSum = 0;
        int counter = 0;

        for (Sheet sheet : roseWork) {

            boolean newRulesPlayers = false;
            
            for (Row row : sheet) {

                String rowStringColor = Tools.getRowStringColor((XSSFRow) row);

                //controllo che non sono a fine rosa squadra
                if (rowStringColor.equals(RoseWorkbook.BLUE_BACK)) break;

                //controllo se sto accedendo ai giocatori del nuovo ordinamento
                if (rowStringColor.equals(RoseWorkbook.YELLOW_BACK)) {
                    newRulesPlayers = true;
                    continue;
                }

                //aggiungo il dq se sono nei giocatori nuovo oridnamento
                if (newRulesPlayers) {

                    int dq = (int) row.getCell(RoseWorkbook.COLUMN_HEADERS.get("DELTA Q")).getNumericCellValue();
                    dqSum += dq;
                    counter += 1;
                }
            }
        }

        setMediaDQRose(dqSum/counter);
    }
    
}
