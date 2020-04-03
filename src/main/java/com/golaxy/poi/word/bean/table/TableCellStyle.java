package com.golaxy.poi.word.bean.table;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

public class TableCellStyle {
    /**
     * 背景颜色
     */
    private String backgroundColor;
    /**
     * 字体大小
     */
    private Integer fontSize;
    /**
     * 字体颜色
     */
    private String fontColor;
    /**
     * 是否加粗
     */
    private boolean isBold=false;
    /**
     * 是否斜体
     */
    private boolean isItalic=false;
    /**
     * 对齐方式
     */
    private ParagraphAlignment textAlign=ParagraphAlignment.LEFT;


    public String getBackgroundColor() {
        return backgroundColor;
    }

    public void setBackgroundColor(String backgroundColor) {
        this.backgroundColor = backgroundColor;
    }

    public Integer getFontSize() {
        return fontSize;
    }

    public void setFontSize(Integer fontSize) {
        this.fontSize = fontSize;
    }

    public String getFontColor() {
        return fontColor;
    }

    public void setFontColor(String fontColor) {
        this.fontColor = fontColor;
    }

    public boolean isBold() {
        return isBold;
    }

    public void setBold(boolean bold) {
        isBold = bold;
    }

    public boolean isItalic() {
        return isItalic;
    }

    public void setItalic(boolean italic) {
        isItalic = italic;
    }

    public ParagraphAlignment getTextAlign() {
        return textAlign;
    }

    public void setTextAlign(ParagraphAlignment textAlign) {
        this.textAlign = textAlign;
    }
}
