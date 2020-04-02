package com.golaxy.poi.word.bean;

public class TitleStyle {
    private int level;
    private String name;
    private int fontSize;

    public TitleStyle(int level, String name, int fontSize) {
        this.level = level;
        this.name = name;
        this.fontSize = fontSize;
    }

    public int getLevel() {
        return level;
    }

    public String getName() {
        return name;
    }

    public int getFontSize() {
        return fontSize;
    }
}
