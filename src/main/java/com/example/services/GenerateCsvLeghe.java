package com.example.services;

import java.io.FileWriter;
import java.io.PrintWriter;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.example.models.QuotationWorkbook;
import com.example.models.RoseWorkbook;
import com.example.utility.Tools;

public class GenerateCsvLeghe {

    private RoseWorkbook rose;
    private QuotationWorkbook quot;
    private String fileName = "data/export.csv";

    
    public GenerateCsvLeghe(RoseWorkbook rose, QuotationWorkbook quot) {

        this.rose = rose;
        this.quot = quot;
    }

    public void generateFile() {

        try (PrintWriter writer = new PrintWriter(new FileWriter(fileName))) {


            writer.println("$");

            sheetsLoop:
            for (Sheet sheet : rose) {

                String teamName = Tools.formatTeamName(sheet.getSheetName());
                playersLoop:
                for (int i = 2; i <= sheet.getLastRowNum(); i++) {

                    //prendo il colore della riga corrente
                    XSSFRow currentRow = (XSSFRow) sheet.getRow(i) ;
                    XSSFColor currentRowColor = currentRow.getCell(0).getCellStyle().getFillForegroundColorColor() != null ? currentRow.getCell(0).getCellStyle().getFillForegroundColorColor() : Tools.stringToColor("#ffffff");
                    String currentRowColorString = Tools.colorToHex(currentRowColor);

                    //se incontro riga gialla continuo il ciclo e switcho oldplayers
                    if (currentRowColorString.equals(RoseWorkbook.YELLOW_BACK)) {
                        continue playersLoop;
                    }

                    //se incontro riga blu cambio sheet
                    if (currentRowColorString.equals(RoseWorkbook.BLUE_BACK)) continue sheetsLoop;

                    Optional<XSSFRow> rowData = checkPlayer(currentRow);

                    if (!rowData.isPresent()) continue playersLoop;

                    String idFantaleghe = String.valueOf((int) rowData.get().getCell(0).getNumericCellValue());
                    int valore = (int) currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("VALORE DI SVINCOLO")).getNumericCellValue();
                    writer.println(teamName + " ," + idFantaleghe + "," + valore);
                }

                writer.println("$");
            } 
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private Optional<XSSFRow> checkPlayer(XSSFRow currentRow) {

        String player_name = currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("NOME GIOCATORE")).getStringCellValue();

        XSSFSheet sheet = this.quot.getSheetAt(0);

        Optional<XSSFRow> rowData = Optional.empty();

        //inizio a scorrere la lista di quotazioni
        for (Row row : sheet) {

            //se arrivo alla fine, finisco il ciclo
            if (row.getCell(0) == null) return rowData;

            String player_name_in_quot = row.getCell(QuotationWorkbook.PLAYER_NAME_COLUMN).getStringCellValue();

            if (player_name.equals(player_name_in_quot)) {

            rowData = Optional.of((XSSFRow) row);

            break;
            }
        }
        return rowData;
    }

}