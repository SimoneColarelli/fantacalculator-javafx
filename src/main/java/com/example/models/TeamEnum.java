package com.example.models;

public enum TeamEnum {
    ATLETICO_ABUSIVO("Atletico Abusivo"),
    ZARRO_TEAM("Zarro Team"),
    RED_DRAGON("Red Dragon"),
    REAL_MADRINK("Real Madrink"),
    SPAL_LETTI("Spal Letti"),
    BAYERN_MUTEN("Bayern Muten"),
    I_CAMMELLONI("I Cammelloni"),
    BOMBERONI("Bomberoni");

    private final String teamName;

    TeamEnum(String teamName) {
        this.teamName = teamName;
    }

    public String getTeamName() {
        return teamName;
    }
}
