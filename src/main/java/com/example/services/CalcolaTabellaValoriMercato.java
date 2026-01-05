package com.example.services;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.example.models.QuotationWorkbook;
import com.example.models.RoseWorkbook;
import com.example.utility.Tools;

public class CalcolaTabellaValoriMercato {
    
    private XSSFWorkbook tabellaMercatoRose;

    public XSSFWorkbook getTabellaMercatoRose() {
        return tabellaMercatoRose;
    }

    public void setTabellaMercatoRose(XSSFWorkbook tabellaMercatoRose) {
        this.tabellaMercatoRose = tabellaMercatoRose;
    }

    private XSSFWorkbook tabellaMercatoQuotazioni;

    public XSSFWorkbook getTabellaMercatoQuotazioni() {
        return tabellaMercatoQuotazioni;
    }

    public void setTabellaMercatoQuotazioni(XSSFWorkbook tabellaMercatoQuotazioni) {
        this.tabellaMercatoQuotazioni = tabellaMercatoQuotazioni;
    }

    public CalcolaTabellaValoriMercato(XSSFWorkbook roseAttuali, XSSFWorkbook quotazioniAttuali, int bilancioMedio) {

        creaTabelle(roseAttuali, quotazioniAttuali);
        riempiTabelle(bilancioMedio);
        underlineGiocNonOccupati(roseAttuali, this.tabellaMercatoQuotazioni);

    }

    private void underlineGiocNonOccupati(XSSFWorkbook roseAttuali, XSSFWorkbook tabMercatoQuot) {
        
        Sheet quotSheet = tabMercatoQuot.getSheetAt(0);

        XSSFCellStyle onListStyle = tabMercatoQuot.createCellStyle();
        XSSFColor color = Tools.stringToColor("#c9daf8");
        onListStyle.setFillForegroundColor(color);
        onListStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle onListPriceStyle = tabMercatoQuot.createCellStyle();
        XSSFColor colorPrice = Tools.stringToColor("#d9ead3");
        onListPriceStyle.setFillForegroundColor(colorPrice);
        onListPriceStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (int i = 2; i <= quotSheet.getLastRowNum(); i++) {

            XSSFRow currentRow = (XSSFRow) quotSheet.getRow(i);

            //fermo il metodo se arrivo alla fine del foglio
            if (currentRow.getCell(0) == null || currentRow.getCell(0).getCellType() == CellType.BLANK) break;

            String currentNome = currentRow.getCell(3).getStringCellValue();

            if (currentNome == null) break;

            boolean freePlayer = true;

            //cerco nella tabella delle squadre
            sheetsLoop:
            for (Sheet sheet : roseAttuali) {

                String colcho = "COLCHONEROS";
                String fenomeni = "FENOMENI_DI_PERIFERIA";

                if (sheet.getSheetName().equals(colcho) || sheet.getSheetName().equals(fenomeni)) continue sheetsLoop;
                //scorro i giocatori
                playersLoop:
                for (int j = 2; j <= sheet.getLastRowNum() - 1; j++) {
                    
                    XSSFRow row = (XSSFRow) sheet.getRow(j);
                    String rowColorString = Tools.getRowStringColor(row);

                    if (rowColorString.equals(RoseWorkbook.YELLOW_BACK)) continue playersLoop;

                    if (rowColorString.equals(RoseWorkbook.BLUE_BACK)) continue sheetsLoop;

                    String name = row.getCell(1).getStringCellValue();
                    String formattedName = Tools.formatName(name);

                    if (currentNome.equals(formattedName)) {
                        freePlayer = false;
                        break sheetsLoop;
                    }
                }
            }

            if (freePlayer) {
                for (int k = 0; k <= 12; k++) {
                    currentRow.getCell(k).setCellStyle(onListStyle);
                }
                currentRow.getCell(13).setCellStyle(onListPriceStyle);
            }
        }

        this.tabellaMercatoQuotazioni = tabMercatoQuot;
    }

    private void creaTabelle(XSSFWorkbook roseAttuali, XSSFWorkbook quotazioniAttuali) {

        RoseWorkbook tabellaMercatoRose = null;
        QuotationWorkbook tabellaMercatoQuotazioni = null;

        try (ByteArrayOutputStream byteArrayOutputStreamRose = new ByteArrayOutputStream();
            ByteArrayOutputStream byteArrayOutputStreamQuot = new ByteArrayOutputStream()) {
            // Scrivi il workbook originale in un ByteArrayOutputStream
            roseAttuali.write(byteArrayOutputStreamRose);
            quotazioniAttuali.write(byteArrayOutputStreamQuot);

            // Leggi i dati dal ByteArrayOutputStream in un ByteArrayInputStream
            try (ByteArrayInputStream byteArrayInputStreamRose = new ByteArrayInputStream(byteArrayOutputStreamRose.toByteArray());
                ByteArrayInputStream byteArrayInputStreamQuot = new ByteArrayInputStream(byteArrayOutputStreamQuot.toByteArray())) {
                // Crea una nuova istanza di RoseWorkbook leggendo dal ByteArrayInputStream
                tabellaMercatoRose = new RoseWorkbook(byteArrayInputStreamRose);
                tabellaMercatoQuotazioni = new QuotationWorkbook(byteArrayInputStreamQuot);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        //aggiungo colonne "valore consigliato", "val. cons. - spesa" alle rose e "val. svin. - val. cons."
        for (Sheet sheet : tabellaMercatoRose) {

            sheet.setColumnWidth(11, (150 * 256) / 7);
            sheet.setColumnWidth(12, (104 * 256) / 7);
            sheet.setColumnWidth(13, (104 * 256) / 7);

            //prendo lo stile delle celle header dalla cella 10 dell'header
            XSSFRow headersRow = (XSSFRow) sheet.getRow(1);
            XSSFCellStyle headerCellStyle = headersRow.getCell(10).getCellStyle();

            //devo clonare questo stile per motivi di overriding
            XSSFCellStyle prezzoConsigliatoStyle = tabellaMercatoRose.createCellStyle();
            prezzoConsigliatoStyle.cloneStyleFrom(headerCellStyle);

            //aggiungo cella valore consigliato nell'header (la creo se non c'è)
            XSSFCell prezzoConsigliato = headersRow.getCell(11);
            if (prezzoConsigliato == null) prezzoConsigliato = headersRow.createCell(11);
            prezzoConsigliato.setCellStyle(prezzoConsigliatoStyle);
            prezzoConsigliato.setCellValue("VALORE CONSIGLIATO");

            //aggiungo cella diff. nell'header (la creo se non c'è)
            XSSFCell differenzaPrezzi = headersRow.getCell(12);
            if (differenzaPrezzi == null) differenzaPrezzi = headersRow.createCell(12);
            differenzaPrezzi.setCellStyle(prezzoConsigliatoStyle);
            differenzaPrezzi.setCellValue("VAL. CONS. - SPESA");

            //aggiungo cella diff.2 nell'header (la creo se non c'è)
            XSSFCell differenzaPrezzi2 = headersRow.getCell(13);
            if (differenzaPrezzi2 == null) differenzaPrezzi2 = headersRow.createCell(13);
            differenzaPrezzi2.setCellStyle(prezzoConsigliatoStyle);
            differenzaPrezzi2.setCellValue("VAL. SVINC. - VAL. CONS.");
        }

        //aggiungo colonna prezzo consigliato alle quotazioni
        Sheet sheet = tabellaMercatoQuotazioni.getSheetAt(0);

        sheet.setColumnWidth(13, (137 * 256) / 7);

        XSSFRow headersQuotRow = (XSSFRow) sheet.getRow(1);
        XSSFCellStyle headerCellStyle = headersQuotRow.getCell(12).getCellStyle();
        XSSFCellStyle prezzoConsigliatoStyle = tabellaMercatoQuotazioni.createCellStyle();
        prezzoConsigliatoStyle.cloneStyleFrom(headerCellStyle);

        XSSFCell prezzoConsigliato = headersQuotRow.getCell(13);
        if (prezzoConsigliato == null) prezzoConsigliato = headersQuotRow.createCell(13);
        prezzoConsigliato.setCellStyle(prezzoConsigliatoStyle);
        prezzoConsigliato.setCellValue("PREZZO CONSIGLIATO");

        this.tabellaMercatoRose = tabellaMercatoRose;
        this.tabellaMercatoQuotazioni = tabellaMercatoQuotazioni;
    }

    private void riempiTabelle(int bilancioMedio) {

        XSSFWorkbook tabellaMercatoRose = this.tabellaMercatoRose;
        XSSFWorkbook tabellaMercatoQuotazioni = this.tabellaMercatoQuotazioni;


        //prendo foglio delle quotazioni
        Sheet sheet = tabellaMercatoQuotazioni.getSheetAt(0);

        for (int j = 2; j <= sheet.getLastRowNum(); j++) {

            XSSFRow currentRow = (XSSFRow) sheet.getRow(j);

            //se arrivo alla fine esco
            if (currentRow.getCell(0) == null || currentRow.getCell(0).getCellType() == CellType.BLANK) break;

            //prendo nome giocatore
            String nomeGiocatore = currentRow.getCell(3).getStringCellValue();

            //prendo fvm consigliato
            double fvm_m = currentRow.getCell(12).getNumericCellValue();

            // calcolo prezzo consigliato
            int prezzoConsigliato = (int) ((fvm_m/1000)*bilancioMedio);

            //costruisco la cella per il prezzo consigliato
            XSSFCell prezzoConsigliatoCellQuot = currentRow.getCell(13);
            if (prezzoConsigliatoCellQuot == null) prezzoConsigliatoCellQuot = currentRow.createCell(13);
            prezzoConsigliatoCellQuot.setCellValue(prezzoConsigliato);

            //inizio a cercare tra le rose
            sheetRoseLoop:
            for (Sheet sheetRose : tabellaMercatoRose) {

                playerLoop:
                for (int i = 2; i <= sheetRose.getLastRowNum() - 1; i++) {
    
                    XSSFRow currentRowRose = (XSSFRow) sheetRose.getRow(i);
    
                    XSSFColor currentColor = currentRowRose.getCell(0).getCellStyle().getFillForegroundColorColor();
    
                    String currentStringColor = currentColor == null ? "#ffffff" : Tools.colorToHex(currentColor);
    
                    if ((currentStringColor.equals(RoseWorkbook.YELLOW_BACK)) || (currentStringColor.equals(RoseWorkbook.RED_BACK))) continue playerLoop;
    
                    if (currentStringColor.equals(RoseWorkbook.BLUE_BACK)) break playerLoop;

                    //vedo nome giocatore
                    String nomeGiocatoreRose = currentRowRose.getCell(RoseWorkbook.COLUMN_HEADERS.get("NOME GIOCATORE")).getStringCellValue();

                    //controllo che non sia capitato per sbaglio in una riga che non è un giocatore
                    if (nomeGiocatoreRose == null || nomeGiocatoreRose.equals("")) continue playerLoop;

                    //formatto il nome
                    String nomeGiocatoreRoseFormat = Tools.formatName(nomeGiocatoreRose);

                    if (nomeGiocatore.equals(nomeGiocatoreRoseFormat)) {
                        
                        //calcolo la differenza di prezzo
                        int val_spesa = (int) currentRowRose.getCell(RoseWorkbook.COLUMN_HEADERS.get("SPESA")).getNumericCellValue();
                        int diff_prezzo = prezzoConsigliato - val_spesa;

                        //calcolo differenza tra val svincolo e val consigliato
                        int val_svinc = (int) currentRowRose.getCell(RoseWorkbook.COLUMN_HEADERS.get("VALORE DI SVINCOLO")).getNumericCellValue();
                        int diff_prezzo_2 = val_svinc - prezzoConsigliato;

                        //creo la cella per aggiungere prezzo consigliato
                        XSSFCell prezzoConsigliatoCellRose = currentRowRose.getCell(11);
                        if (prezzoConsigliatoCellRose == null) prezzoConsigliatoCellRose = currentRowRose.createCell(11);
                        prezzoConsigliatoCellRose.setCellValue(prezzoConsigliato);

                        //creo la cella per aggiungere diff prezzo
                        XSSFCell diffPrezzoCellRose = currentRowRose.getCell(12);
                        if (diffPrezzoCellRose == null) diffPrezzoCellRose = currentRowRose.createCell(12);
                        diffPrezzoCellRose.setCellValue(diff_prezzo);

                        //creo la cella epr aggiungere diff prezzo 2
                        XSSFCell diffPrezzoCellRose2 = currentRowRose.getCell(13);
                        if (diffPrezzoCellRose2 == null) diffPrezzoCellRose2 = currentRowRose.createCell(13);
                        diffPrezzoCellRose2.setCellValue(diff_prezzo_2);

                        //definisco il colore del valore consigliato
                        byte[] rgb = Tools.hexToRgb("#c9daf8");
                        XSSFColor b_color = new XSSFColor(rgb, null);

                        //assegno il colore al prezzo consigliato
                        XSSFCellStyle prezzoConsigliatoCellRoseStyle = tabellaMercatoRose.createCellStyle();
                        prezzoConsigliatoCellRoseStyle.cloneStyleFrom(currentRowRose.getCell(2).getCellStyle());
                        prezzoConsigliatoCellRoseStyle.setFillForegroundColor(b_color);
                        prezzoConsigliatoCellRoseStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        prezzoConsigliatoCellRose.setCellStyle(prezzoConsigliatoCellRoseStyle);

                        //definisco come dare il colore alla differenza valore
                        byte[] rgb_1 = Tools.hexToRgb("#d9ead3");
                        XSSFColor gr_color = new XSSFColor(rgb_1, null);
                        byte[] rgb_2 = Tools.hexToRgb("#ea9999");
                        XSSFColor red_color = new XSSFColor(rgb_2, null);

                        //creo gli stili per diff prezzo e diff prezzo 2
                        XSSFCellStyle diffPrezzoCellRoseStyle = tabellaMercatoRose.createCellStyle();
                        XSSFCellStyle diffPrezzoCellRoseStyle2 = tabellaMercatoRose.createCellStyle();
                        diffPrezzoCellRoseStyle.cloneStyleFrom(currentRowRose.getCell(2).getCellStyle());

                        diffPrezzoCellRoseStyle.setFillForegroundColor(gr_color);
                        diffPrezzoCellRoseStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                        diffPrezzoCellRoseStyle2.cloneStyleFrom(diffPrezzoCellRoseStyle);

                        //determino se contenuto di diff prezzo è positivo o negativo
                        if (diffPrezzoCellRose.getNumericCellValue() < 0) {
                            diffPrezzoCellRoseStyle.setFillForegroundColor(red_color);
                            diffPrezzoCellRoseStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        }

                        diffPrezzoCellRose.setCellStyle(diffPrezzoCellRoseStyle);


                        //determino se contenuto di diff prezzo 2 è positivo o negativo
                        if (diffPrezzoCellRose2.getNumericCellValue() < 0) {
                            diffPrezzoCellRoseStyle2.setFillForegroundColor(red_color);
                            diffPrezzoCellRoseStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        }

                        diffPrezzoCellRose2.setCellStyle(diffPrezzoCellRoseStyle2);

                        break sheetRoseLoop;
                    }
                }
            }
        }
    }
}
