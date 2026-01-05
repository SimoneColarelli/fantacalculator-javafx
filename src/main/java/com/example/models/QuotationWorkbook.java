package com.example.models;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Quotation_workbook
 */
public class QuotationWorkbook extends XSSFWorkbook {

    public static final int QUOT_COLUMN = 8;

    public static final int INIT_QUOT_COLUMN = 9;

    public static final int FIRST_PLAYER_ROW = 2;

    public static final int PLAYER_NAME_COLUMN = 3;

    public QuotationWorkbook(InputStream is) throws IOException {
        super(is);
    }
}