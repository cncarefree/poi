package com.golaxy.poi.word;

import com.golaxy.poi.word.bean.table.Table;
import com.golaxy.poi.word.bean.table.TableCell;
import com.golaxy.poi.word.bean.table.TableCellStyle;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TableRowAlign;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class DocumentBuilderTest {
    @Test
    public void testCreateTitle() throws IOException {
        DocumentBuilder builder = new DocumentBuilder();
        builder.setHeader(HeaderFooterType.DEFAULT, ParagraphAlignment.CENTER,"1212121212");
        builder.setFooter(HeaderFooterType.DEFAULT,ParagraphAlignment.RIGHT,"ssdsds");
        builder.createTitle(1,"一、测试测试");
        builder.createTitle(2,"1.1 测试测试");
        builder.createTitle(3,"1.1.1 测试测试");
        builder.createTitle(4,"1.1.1.1 测试测试");
        List<TableCell> head=new ArrayList<TableCell>();
        TableCellStyle headStyle=new TableCellStyle();
        headStyle.setBackgroundColor("CCCCCC");
        headStyle.setBold(true);
        headStyle.setTextAlign(ParagraphAlignment.RIGHT);
        head.add(new TableCell("序号",headStyle,22));
        head.add(new TableCell("姓名",headStyle,44));
        head.add(new TableCell("性别",headStyle,22));
        head.add(new TableCell("联系方式",headStyle,110));
        head.add(new TableCell("地址",headStyle,200));
        Table table=new Table(null, TableRowAlign.CENTER,head);
        for (int i = 0; i < 50; i++) {
            List<String> list=new ArrayList<>();
            list.add(i+1+"");
            list.add("张三"+i);
            list.add("男");
            list.add(13000000000L+i+"");
            list.add("北京市");
            table.appendRow(list);
        }
       builder.createTable(table);
        builder.build(new FileOutputStream("/Users/jiangzhaoyue/Downloads/poiTest/1.docx"));
    }



}
