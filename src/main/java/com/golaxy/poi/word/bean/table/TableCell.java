package com.golaxy.poi.word.bean.table;

public class TableCell {
    private String text;
    private TableCellStyle style;
    private Integer width;

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public TableCellStyle getStyle() {
        return style;
    }

    public void setStyle(TableCellStyle style) {
        this.style = style;
    }

    public Integer getWidth() {
        return width;
    }

    public void setWidth(Integer width) {
        this.width = width;
    }

    public TableCell(String text, TableCellStyle style) {
        this.text = text;
        this.style = style;
    }

    public TableCell(String text) {
        this.text = text;
    }

    public TableCell(String text, TableCellStyle style, Integer width) {
        this.text = text;
        this.style = style;
        this.width = width;
    }

    public TableCell(String text, Integer width) {
        this.text = text;
        this.width = width;
    }

}
