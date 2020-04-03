package com.golaxy.poi.word.bean.table;

import org.apache.poi.xwpf.usermodel.TableRowAlign;

import java.util.ArrayList;
import java.util.List;

public class Table {
    /**
     * 表格宽度
     */
    private Integer width;
    /**
     * 表格整体居中
     */
    private TableRowAlign tableAlign=TableRowAlign.CENTER;
    /**
     * 表头
     */
    private List<TableCell> header;
    /**
     * 行
     */
    private List<TableRow> rows=new ArrayList<TableRow>();



    /**
     * 创建表格
     * @param width 设置宽度，不设置则按填充内容适配，单位:磅,五号字为10.5磅
     * @param tableAlign 表格对齐（相对全文，默认居中）
     * @param header    设置表头
     */
    public Table(Integer width, TableRowAlign tableAlign, List<TableCell> header) {
        this.width = width;
        this.tableAlign = tableAlign;
        this.header = header;
    }

    /**
     * 设置表格宽度
     * @param width 设置宽度，不设置则按填充内容适配，单位:磅,五号字为10.5磅
     */
    public void setWidth(Integer width) {
        this.width = width;
    }

    /**
     * 对齐方式
     * @param tableAlign 表格对齐（相对全文，默认居中）
     */
    public void setTableAlign(TableRowAlign tableAlign) {
        this.tableAlign = tableAlign;
    }

    /**
     * 设置表头
     * @param header
     */
    public void setHeader(List<TableCell> header) {
        this.header = header;
    }
    /**
     * 添加行
     * @param rowCell 行里头的每个字段
     */
    public void appendRowByTabelCell(List<TableCell> rowCell){
        rows.add(new TableRow(rowCell));
    }
    /**
     * 添加行
     * @param textList 纯文本行
     */
    public void appendRow(List<String> textList){
        List<TableCell> list=new ArrayList<>();
        textList.forEach(text->{
            list.add(new TableCell(text));
        });
        rows.add(new TableRow(list));
    }

    public Integer getWidth() {
        return width;
    }

    public TableRowAlign getTableAlign() {
        return tableAlign;
    }

    public List<TableCell> getHeader() {
        return header;
    }

    public List<TableRow> getRows() {
        return rows;
    }

    @Override
    public String toString() {
        return "Table{" +
                "width=" + width +
                ", tableAlign=" + tableAlign +
                ", header=" + header +
                ", rows=" + rows +
                '}';
    }

}
