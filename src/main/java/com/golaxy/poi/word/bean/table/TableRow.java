package com.golaxy.poi.word.bean.table;

import java.util.List;

public class TableRow {
    private List<TableCell> cell;

    protected TableRow(List<TableCell> cell) {
        this.cell = cell;
    }


    public List<TableCell> getCell() {
        return cell;
    }

    protected void setCell(List<TableCell> cell) {
        this.cell = cell;
    }
}
