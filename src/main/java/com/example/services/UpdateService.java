package com.example.services;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.example.models.QuotationWorkbook;
import com.example.models.RoseWorkbook;
import com.example.utility.Tools;

public class UpdateService {

    private QuotationWorkbook nuoveQuot;

    private RoseWorkbook updatedRose;

    private ArrayList<String> alertedPlayersList;

    public ArrayList<String> getAlertedPlayersList() {
        return alertedPlayersList;
    }

    public UpdateService(RoseWorkbook rose, QuotationWorkbook nuoveQuot) {

        this.nuoveQuot = nuoveQuot;
        this.updatedRose = rose;
    }

    public RoseWorkbook getUpdatedRose() {
        return updatedRose;
    }

    public QuotationWorkbook getNuoveQuot() {
        return nuoveQuot;
    }

    public void doCompleteUpdate() {

        RoseWorkbook rose = this.updatedRose;

        //definisco il colore da dare a quelli che non trovo e la lista di stringhe a cui aggiungere
        String alertColor = "#ffe599";
        ArrayList<String> alertedPlayers = new ArrayList<>();

        //inizializzo workbook rose aggiornate
        RoseWorkbook roseAggiornate = null;

        try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
            // Scrivi il workbook originale in un ByteArrayOutputStream
            rose.write(byteArrayOutputStream);

            // Leggi i dati dal ByteArrayOutputStream in un ByteArrayInputStream
            try (ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray())) {
                // Crea una nuova istanza di RoseWorkbook leggendo dal ByteArrayInputStream
                roseAggiornate = new RoseWorkbook(byteArrayInputStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        sheetsLoop:
        for (Sheet sheet : roseAggiornate) {

            //inizializzo la booleana che mi dice se sono tra i gicoatori vecchio ordinamento o no
            boolean oldPlayers = true;

            playersLoop:
            for (int i = 2; i <= sheet.getLastRowNum(); i++) {

                //prendo il colore della riga corrente
                XSSFRow currentRow = (XSSFRow) sheet.getRow(i) ;
                XSSFColor currentRowColor = currentRow.getCell(0).getCellStyle().getFillForegroundColorColor() != null ? currentRow.getCell(0).getCellStyle().getFillForegroundColorColor() : Tools.stringToColor("#ffffff");
                String currentRowColorString = Tools.colorToHex(currentRowColor);

                //se incontro riga gialla continuo il ciclo e switcho oldplayers
                if (currentRowColorString.equals(RoseWorkbook.YELLOW_BACK)) {
                    oldPlayers = false;
                    continue playersLoop;
                }

                //se incontro riga blu cambio sheet
                if (currentRowColorString.equals(RoseWorkbook.BLUE_BACK)) continue sheetsLoop;

                System.out.println("nome giocatore: " + currentRow.getCell(1).getStringCellValue());

                Optional<XSSFRow> rowData = checkPlayer(currentRow);

                //controllo se è vecchio ordinamento
                if (oldPlayers) {

                    //se non è in prestito, calcolo deprezzamento
                    if (!currentRowColorString.equals(RoseWorkbook.GREEN_BACK)) calculateOldRulesValue(currentRow);

                    //controllo se non è in serie a e nel caso cambio colore
                    if (!rowData.isPresent()) {
                        for (int k = 0; k <= 10; k++) {
                            XSSFCellStyle currentStyle = currentRow.getCell(1).getCellStyle();
                            XSSFCellStyle cellStyle = (XSSFCellStyle) sheet.getWorkbook().createCellStyle();
                            cellStyle.cloneStyleFrom(currentStyle);
                            cellStyle.setFillForegroundColor(Tools.stringToColor(alertColor));
                            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            currentRow.getCell(k).setCellStyle(cellStyle);
                        }

                        //se non c'è nella lista lo aggiungo nella lista degli alerted players
                        String nomeGiocatore = currentRow.getCell(1).getStringCellValue();
                        alertedPlayers.add(nomeGiocatore);
                    }
                }

                //altrimenti (se è nuovo ordinamento) calcolo valore a seconda che sia in prestito o meno
                else {

                    //se non è presente lo metto negli alerted
                    if (!rowData.isPresent()) {
                        for (int k = 0; k <= 10; k++) {
                            XSSFCellStyle currentStyle = currentRow.getCell(1).getCellStyle();
                            XSSFCellStyle cellStyle = (XSSFCellStyle) sheet.getWorkbook().createCellStyle();
                            cellStyle.cloneStyleFrom(currentStyle);
                            cellStyle.setFillForegroundColor(Tools.stringToColor(alertColor));
                            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            currentRow.getCell(k).setCellStyle(cellStyle);
                        }

                        //se non c'è nella lista lo aggiungo nella lista degli alerted players
                        String nomeGiocatore = currentRow.getCell(1).getStringCellValue();
                        alertedPlayers.add(nomeGiocatore);
                    }

                    else {

                        //calcola il valore (il metodo tiene conto se è in prestito o meno)
                        updatePlayerWithNewRules(currentRow, rowData.get());
                    }
                }
            }
        }

        this.updatedRose = roseAggiornate;

        this.alertedPlayersList = alertedPlayers;
    }


    private void updatePlayerWithNewRules(XSSFRow currentRow, XSSFRow quotationRow) {

        XSSFColor currentColor = currentRow.getCell(1).getCellStyle().getFillForegroundColorColor();
        
        String currentStringColor = currentColor != null ? Tools.colorToHex(currentColor) : "#ffffff";

        System.out.println("colore riga è: " + currentStringColor);

        int new_current_value = (int) currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("VALORE DI SVINCOLO")).getNumericCellValue();
        int new_quotation = (int) quotationRow.getCell(QuotationWorkbook.QUOT_COLUMN).getNumericCellValue();
        int new_delta = (int) currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("DELTA Q")).getNumericCellValue();

        int current_quotaz = (int) currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("QUOTAZ")).getNumericCellValue();

        int delta = new_quotation - current_quotaz;

        if (currentStringColor.equals(RoseWorkbook.GREEN_BACK)) {
            if ((new_quotation-current_quotaz) < 0) delta = 0;
        }

        //calcolo nuovo delta q
        new_delta += delta;

        int initial_value = (int) currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("VALORE INIZIALE")).getNumericCellValue();

        //calcolo nuovo valore di svincolo giocatore
        new_current_value = this.calculateCurrentValue(initial_value, new_delta);

        //se non era in prestito metti lo sfondo bianco
        if (!(currentStringColor.equals(RoseWorkbook.GREEN_BACK))) {

            int current_index = 0;
            for (Cell current_cell : currentRow) {

                XSSFCellStyle currentStyle = (XSSFCellStyle) current_cell.getCellStyle();
                XSSFCellStyle cellStyle = (XSSFCellStyle) current_cell.getSheet().getWorkbook().createCellStyle();
                cellStyle.cloneStyleFrom(currentStyle);

                if (current_index == RoseWorkbook.COLUMN_HEADERS.get("VALORE DI SVINCOLO")) {
                    XSSFColor svincoloColor = Tools.stringToColor("#ddebf7");
                    cellStyle.setFillForegroundColor(svincoloColor);
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    current_cell.setCellStyle(cellStyle);
                    current_index += 1;
                    continue;
                }

                XSSFColor whiteColor = Tools.stringToColor("#ffffff");
                cellStyle.setFillForegroundColor(whiteColor);
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                current_cell.setCellStyle(cellStyle);

                if (current_index > 10) break;

                current_index += 1;
            }
        }

        //una volta calcolati i nuovi valori, aggiorna il giocatore
        currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("QUOTAZ")).setCellValue(new_quotation);
        currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("DELTA Q")).setCellValue(new_delta);
        currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("VALORE DI SVINCOLO")).setCellValue(new_current_value);
    }

    private int calculateCurrentValue(int initial_value, int new_delta) {
        int new_current_value = initial_value;
        int delta_sign = Tools.signNum(new_delta);
        int delta_abs = new_delta*delta_sign;

        if (new_delta == 0) return new_current_value;

        for (int i = delta_abs; i >=1; i--) {

            if (new_current_value >= 1 && new_current_value <= 49) {
                if (delta_sign == -1) {
                    new_current_value = new_current_value - 3*delta_abs;
                    break;
                }
                else new_current_value = (int) ((double) new_current_value + 21.5);
            }

            else if (new_current_value >= 50 && new_current_value <= 99) {
                if (delta_sign == -1) {
                    new_current_value = new_current_value - 8*delta_abs;
                    break;
                }
                else new_current_value = new_current_value + 18;
            }

            else if (new_current_value >= 100 && new_current_value <= 199) {
                if (delta_sign == -1) {
                    new_current_value = new_current_value - 12*delta_abs;
                    break;
                }
                else new_current_value = new_current_value + 12;
            }

            else if (new_current_value >= 200 && new_current_value <= 399) {
                if (delta_sign == -1) {
                    new_current_value = new_current_value - 18*delta_abs;
                    break;
                }
                else new_current_value = new_current_value + 8;
            }

            else if (new_current_value >= 400 && new_current_value <= 599) {
                if (delta_sign == -1) {
                    new_current_value = (int) ((double) new_current_value - 21.5*delta_abs);
                    break;
                }
                else new_current_value = new_current_value + 3;
            }

            else if (new_current_value >= 600 && new_current_value <= 99999) {
                if (delta_sign == -1) {
                    new_current_value = new_current_value - 30*delta_abs;
                    break;
                }
                else new_current_value = new_current_value + 1;
            }

        }

        return new_current_value > 0 ? new_current_value : 1;
    }

    private void calculateOldRulesValue(XSSFRow currentRow) {

        int deprezz = (int) currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("DEPREZZ")).getNumericCellValue();
        int valore_attuale = (int) currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("VALORE DI SVINCOLO")).getNumericCellValue();

        valore_attuale = valore_attuale - deprezz;

        currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("VALORE DI SVINCOLO")).setCellValue(valore_attuale > 0 ? valore_attuale : 1);

    }

    private Optional<XSSFRow> checkPlayer(XSSFRow currentRow) {

        String player_name = currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("NOME GIOCATORE")).getStringCellValue();
        String formatted_player_name = Tools.formatName(player_name);

        XSSFSheet sheet = this.nuoveQuot.getSheetAt(0);

        Optional<XSSFRow> rowData = Optional.empty();

        //inizio a scorrere la lista di quotazioni
        for (Row row : sheet) {

            //se arrivo alla fine, finisco il ciclo
            if (row.getCell(0) == null) return rowData;

            String player_name_in_quot = row.getCell(QuotationWorkbook.PLAYER_NAME_COLUMN).getStringCellValue();

            if (formatted_player_name.equals(player_name_in_quot)) {

            rowData = Optional.of((XSSFRow) row);

            break;
            }
        }

        return rowData;
    }

    public void doQuotUpdate() {

        RoseWorkbook rose = this.updatedRose;

        //definisco il colore da dare a quelli che non trovo e la lista di stringhe a cui aggiungere
        String alertColor = "#ffe599";

        ArrayList<String> alertedPlayers = new ArrayList<>();

        RoseWorkbook roseAggiornate = null;

        try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
            // Scrivi il workbook originale in un ByteArrayOutputStream
            rose.write(byteArrayOutputStream);

            // Leggi i dati dal ByteArrayOutputStream in un ByteArrayInputStream
            try (ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray())) {
                // Crea una nuova istanza di RoseWorkbook leggendo dal ByteArrayInputStream
                roseAggiornate = new RoseWorkbook(byteArrayInputStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        for (Sheet sheet : roseAggiornate) {

            //inizializzo la booleana che mi dice se sono tra i gicoatori vecchio ordinamento o no
            boolean oldPlayers = true;

            //Inizio a scorrere i giocatori 
            playerLoop:
            for (int currentRowNum = RoseWorkbook.FIRST_PLAYER_ROW; currentRowNum <= sheet.getLastRowNum() - 1; currentRowNum++) {

                XSSFRow currentRow = (XSSFRow) sheet.getRow(currentRowNum);

                XSSFColor currentColor = currentRow.getCell(0).getCellStyle().getFillForegroundColorColor();

                String currentStringColor = currentColor == null ? "#ffffff" : Tools.colorToHex(currentColor);

                //se incontro riga gialla vado avanti e cambio da veccchio a nuovo ordinamento
                if ((currentStringColor.equals(RoseWorkbook.YELLOW_BACK))) {

                    oldPlayers = false;
                    continue;
                }

                //se incontro linea blu, cambio sheet
                if (currentStringColor.equals(RoseWorkbook.BLUE_BACK)) {
                    break playerLoop;
                }

                Optional<XSSFRow> rowData = checkPlayer(currentRow);

                //se non c'è nella lista quotazioni lo segnalo di un colore di alert e aggiungo alla lista di giocatori da segnalare
                if (!rowData.isPresent()) {

                    alertedPlayers.add(currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("NOME GIOCATORE")).getStringCellValue());

                    for (int i = 0; i <= 10; i++) {

                        if (i == 8) continue;
                        
                        XSSFCell currentCell = currentRow.getCell(i);
                        // Create a new cell style
                        XSSFCellStyle newCellStyle = (XSSFCellStyle) roseAggiornate.createCellStyle();
                        
                        // Clone the existing style into the new style
                        newCellStyle.cloneStyleFrom(currentCell.getCellStyle());

                        // Create the alert color
                        XSSFColor alert_color = new XSSFColor(Tools.hexToRgb(alertColor), null);

                        // Set the new style's fill color and pattern
                        newCellStyle.setFillForegroundColor(alert_color);
                        newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                        // Apply the new style to the cell
                        currentCell.setCellStyle(newCellStyle);
                    }
                }

                //cosa faccio se è presente
                else {

                    if (!oldPlayers) {
                        XSSFRow quotationRow = checkPlayer(currentRow).get();
                        int new_quot = (int) quotationRow.getCell(9).getNumericCellValue();

                        XSSFCell quotazCell = currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("QUOTAZ"));
                        quotazCell.setCellValue(new_quot);
                    }

                    for (int i = 0; i <= 10; i++) {

                        if (i == 8) continue;
                        
                        XSSFCell currentCell = currentRow.getCell(i);
                        // Create a new cell style
                        XSSFCellStyle newCellStyle = (XSSFCellStyle) roseAggiornate.createCellStyle();
                        
                        // Clone the existing style into the new style
                        newCellStyle.cloneStyleFrom(currentCell.getCellStyle());

                        // Create the alert color
                        XSSFColor white_color = new XSSFColor(Tools.hexToRgb("#ffffff"), null);

                        // Set the new style's fill color and pattern
                        newCellStyle.setFillForegroundColor(white_color);
                        newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                        // Apply the new style to the cell
                        currentCell.setCellStyle(newCellStyle);
                    }


                }
            }
        }

        this.updatedRose = roseAggiornate;

        this.alertedPlayersList = alertedPlayers;
    }

}
