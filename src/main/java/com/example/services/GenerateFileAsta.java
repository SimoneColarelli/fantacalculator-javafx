package com.example.services;

import java.util.Optional;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.example.models.QuotationWorkbook;
import com.example.models.RoseWorkbook;
import com.example.utility.Tools;

public class GenerateFileAsta {
    
    private XSSFWorkbook excelAsta = null;

    private XSSFWorkbook excelRose;

    private XSSFWorkbook excelQuot;


    public GenerateFileAsta(XSSFWorkbook roseAtt, XSSFWorkbook listoneAtt) {
        this.excelRose = roseAtt;
        this.excelQuot = listoneAtt;

        fillRose();

    }

    private Optional<XSSFRow> checkPlayer(XSSFRow currentRow) {

        String player_name = currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("NOME GIOCATORE")).getStringCellValue();
        String formatted_player_name = Tools.formatName(player_name);

        XSSFSheet sheet = this.excelQuot.getSheetAt(0);

        Optional<XSSFRow> rowData = Optional.empty();

        //inizio a scorrere la lista di quotazioni
        for (Row row : sheet) {

            //se arrivo alla fine, finisco il ciclo
            if (row.getCell(0) == null || row.getCell(0).getCellType() == CellType.BLANK) return rowData;
            
            String player_name_in_quot = row.getCell(QuotationWorkbook.PLAYER_NAME_COLUMN).getStringCellValue();

            if (formatted_player_name.equals(player_name_in_quot)) {

            rowData = Optional.of((XSSFRow) row);

            break;
            }
        }

        return rowData;
    }


    private void fillRose() {

        XSSFWorkbook excelAsta = new XSSFWorkbook();

        Sheet mainSheet = excelAsta.createSheet();

        sheetLoop:
        for (Sheet sheet : this.excelRose) {

            int lastRowNum = mainSheet.getLastRowNum();

            XSSFRow teamNameRow = (XSSFRow) mainSheet.createRow(lastRowNum + 1);

            String teamName = sheet.getSheetName();

            System.out.println(teamName);

            teamNameRow.createCell(0).setCellValue(teamName);

            playersLoop:
            for (int i = 2; i <= sheet.getLastRowNum(); i++) {

            XSSFRow currentRow = (XSSFRow) sheet.getRow(i) ;
            XSSFColor currentRowColor = currentRow.getCell(0).getCellStyle().getFillForegroundColorColor() != null ? currentRow.getCell(0).getCellStyle().getFillForegroundColorColor() : Tools.stringToColor("#ffffff");
            String currentRowColorString = Tools.colorToHex(currentRowColor);

            if (currentRowColorString.equals(RoseWorkbook.YELLOW_BACK)) {
                continue playersLoop;
            }

            //se incontro riga blu cambio sheet
            if (currentRowColorString.equals(RoseWorkbook.BLUE_BACK)) continue sheetLoop;

            Optional<XSSFRow> rowData = checkPlayer(currentRow);

            if (!rowData.isPresent()) continue playersLoop;

            XSSFRow quotRow = rowData.get();

            String ruoli = quotRow.getCell(2).getStringCellValue();
            String nome = quotRow.getCell(3).getStringCellValue();
            String squadra = quotRow.getCell(4).getStringCellValue().substring(0, 3);
            String costo = String.valueOf(currentRow.getCell(2).getNumericCellValue());

            lastRowNum = lastRowNum + 1;

            XSSFRow newRow = (XSSFRow) mainSheet.createRow(lastRowNum);

            newRow.createCell(0).setCellValue(ruoli);
            newRow.createCell(1).setCellValue(nome);
            newRow.createCell(2).setCellValue(squadra);
            newRow.createCell(3).setCellValue(costo);
            }
        }

        this.excelAsta = excelAsta;
    }


    //getters and setters
    public XSSFWorkbook getExcelRose() {
        return excelRose;
    }

    public void setExcelRose(XSSFWorkbook excelRose) {
        this.excelRose = excelRose;
    }

    
    public XSSFWorkbook getExcelQuot() {
        return excelQuot;
    }

    public void setExcelQuot(XSSFWorkbook excelQuot) {
        this.excelQuot = excelQuot;
    }

    public XSSFWorkbook getExcelAsta() {
        return excelAsta;
    }

    public void setExcelAsta(XSSFWorkbook excelAsta) {
        this.excelAsta = excelAsta;
    }
}

