package com.example.services;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.example.models.RoseWorkbook;
import com.example.utility.Tools;

public class CalcolaMediaPatrimoni {
    
    private int mediaPatrimoni;

    public CalcolaMediaPatrimoni (File fileBilanci, File fileRose) throws IOException {
        calcolaMedia(fileBilanci, fileRose);
    }
    
    public int getMediaPatrimoni() {
        return mediaPatrimoni;
    }

    public void setMediaPatrimoni(int mediaPatrimoni) {
        this.mediaPatrimoni = mediaPatrimoni;
    }

    private void calcolaMedia(File fileBilanci, File fileRose) throws IOException {
        try (
            FileInputStream bilanciInputStream = new FileInputStream(fileBilanci);
            FileInputStream roseInputStream = new FileInputStream(fileRose);
            XSSFWorkbook bilanciWork = new XSSFWorkbook(bilanciInputStream);
            XSSFWorkbook roseWork = new XSSFWorkbook(roseInputStream);
        ) {

            //sommo i bilanci
            int sum = 0;
            XSSFRow bilanciRow = bilanciWork.getSheetAt(0).getRow(2);
            for (Cell cell : bilanciRow) {
                int cellValue = (int) cell.getNumericCellValue();
                sum += cellValue;
            }
            this.mediaPatrimoni += sum;

            //vado a sommare i valori di svincolo dei giocatori
            for (Sheet sheet : roseWork) {
                
                if (sheet.getSheetName().equals("I_FENOMENI_DI_PERIFERIA") || sheet.getSheetName().equals("COLCHONEROS")) continue;

                playerLoop:
                for (int i = RoseWorkbook.FIRST_PLAYER_ROW; i <= sheet.getLastRowNum() - 1; i++) {

                    XSSFRow currentRow = (XSSFRow) sheet.getRow(i);

                    XSSFColor currentColor = currentRow.getCell(0).getCellStyle().getFillForegroundColorColor();

                    String currentStringColor = currentColor == null ? "#ffffff" : Tools.colorToHex(currentColor);

                    //controllo che sia gialla
                    if ((currentStringColor.equals(RoseWorkbook.YELLOW_BACK))) continue;

                    //controllo che sia blu
                    if (currentStringColor.equals(RoseWorkbook.BLUE_BACK)) break playerLoop;

                    XSSFCell svincoloCell = currentRow.getCell(RoseWorkbook.COLUMN_HEADERS.get("VALORE DI SVINCOLO"));
                    int svincoloValue = (int) svincoloCell.getNumericCellValue();

                    sum += svincoloValue;
                }
            }

            this.mediaPatrimoni = sum / 8;

        } catch(IOException e) {
            e.printStackTrace();
        }
    }
}
