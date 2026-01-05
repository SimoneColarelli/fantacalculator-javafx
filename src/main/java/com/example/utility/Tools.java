package com.example.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.example.models.QuotationWorkbook;
import com.example.models.RoseWorkbook;

public class Tools {
    
    public static byte[] hexToRgb(String hex) {

        hex = hex.replaceAll("[^0-9A-Fa-f]", "");

        return new byte[] {
            (byte) Integer.parseInt(hex.substring(0, 2), 16),
            (byte) Integer.parseInt(hex.substring(2, 4), 16),
            (byte) Integer.parseInt(hex.substring(4, 6), 16)
        };
    }

    public static XSSFColor stringToColor(String colorString) {
        return new XSSFColor(hexToRgb(colorString), null);
    }

    public static String colorToHex (XSSFColor color) {

        byte[] rgb = color.getRGB();

        if (rgb != null) {
            int red = (rgb[0] & 0xFF);
            int green = (rgb[1] & 0xFF);
            int blue = (rgb[2] & 0xFF);

            String hexColor = String.format("#%02X%02X%02X", red, green, blue);
            return hexColor.toLowerCase();
        } else {
            return "#ffffff";
        }
    }

    public static int signNum(int n) {
        if (n < 0) return -1;
        if (n > 0) return 1;
        else return 0;
    }

    public static String formatName(String name) {
        String lowerCaseName = name.toLowerCase();
        char[] charArray = lowerCaseName.toCharArray();
        charArray[0] = Character.toUpperCase(charArray[0]);
        for (int i = 1; i < charArray.length - 2; i++) {
            if (charArray[i] == ' ' || charArray[i] == '-' || charArray[i] == '\'') charArray[i + 1] = Character.toUpperCase(charArray[i + 1]);
            if (charArray[i-1] == 'M' && charArray[i] == 'c') charArray[i+1] = Character.toUpperCase(charArray[i + 1]);

        }

        /*controllo l'ultima lettera se è accentata con l'apice
        if (charArray[charArray.length - 1] == '\'') {
            char [] cuttedArray = Arrays.copyOfRange(charArray, 0, charArray.length - 1);
            int l = cuttedArray.length - 1;
            if (cuttedArray[l] == 'a') cuttedArray[l] = 'à';
            if (cuttedArray[l] == 'e') cuttedArray[l] = 'è';
            if (cuttedArray[l] == 'i') cuttedArray[l] = 'ì';
            if (cuttedArray[l] == 'o') cuttedArray[l] = 'ò';
            if (cuttedArray[l] == 'u') cuttedArray[l] = 'ù';
            formattedName = new String(cuttedArray);
        }

        else {
            formattedName = new String(charArray);
        }*/

        return new String(charArray);
    }

    public static String getRowStringColor(XSSFRow row) {
        
        XSSFColor rowColor = (XSSFColor) row.getCell(0).getCellStyle().getFillForegroundColorColor();
        String rowStringColor = rowColor != null ? Tools.colorToHex(rowColor) : "#ffffff";

        return rowStringColor;
    }

    public static XSSFWorkbook fileToWorkbook(File file) {

        XSSFWorkbook workbook = null;
        try (FileInputStream fis = new FileInputStream(file)) {
            workbook = new XSSFWorkbook(fis);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }

    public static RoseWorkbook fileToRoseWorkbook(File file) {

        RoseWorkbook workbook = null;
        try (FileInputStream fis = new FileInputStream(file)) {
            workbook = new RoseWorkbook(fis);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }

    public static QuotationWorkbook fileToQuotationWorkbook(File file) {

        QuotationWorkbook workbook = null;
        try (FileInputStream fis = new FileInputStream(file)) {
            workbook = new QuotationWorkbook(fis);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }

    public static String formatTeamName(String input) {
        if (input == null || input.isEmpty()) {
            return input;
        }

        // tutto in minuscolo
        String lower = input.toLowerCase();

        // split sugli underscore
        String[] parts = lower.split("_");

        // capitalizza la prima lettera di ogni parte
        StringBuilder sb = new StringBuilder();
        for (String part : parts) {
            if (part.isEmpty()) continue;
            sb.append(Character.toUpperCase(part.charAt(0)))
              .append(part.substring(1))
              .append(" ");
        }

        // rimuove lo spazio finale
        return sb.toString().trim();
    }

}
