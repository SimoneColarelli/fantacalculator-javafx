package com.example.models;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class RoseWorkbook extends XSSFWorkbook{
    
    //Definsico le costanti di questo tipo di workbook

    public static final int TEAM_NAME_ROW = 0;

    public static final int HEADER_ROW = 1;

    public static final int FIRST_PLAYER_ROW = 2;

    public static final Map<String, Integer> COLUMN_HEADERS = Map.ofEntries(
        Map.entry("ID", 0),
        Map.entry("NOME GIOCATORE", 1),
        Map.entry("SPESA", 2),
        Map.entry("DATA", 3),
        Map.entry("FASCIA", 4),
        Map.entry("VALORE INIZIALE", 5),
        Map.entry("QUOTAZ", 6),
        Map.entry("DELTA Q", 7),
        Map.entry("VALORE DI SVINCOLO", 8),
        Map.entry("DEPREZZ", 9),
        Map.entry("SCADENZA CONTRATTO", 10)
    );


    //Definisco i colori costanti di questo tipo di workbook
    public static final String WHITE_BACK = "#ffffff";

    public static final String YELLOW_BACK = "#ffff00";

    public static final String GREEN_BACK = "#a9d08e";

    public static final String RED_BACK = "#f4b084";

    public static final String BLUE_BACK = "#8ea9db";

    public RoseWorkbook(InputStream is) throws IOException {
        super(is);
    }
}
