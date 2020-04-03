package com.golaxy.poi.word;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

public class DocumentBuilderTest {
    @Test
    public void testCreateTitle() throws IOException {
        DocumentBuilder builder = new DocumentBuilder();
        builder.setHeader(HeaderFooterType.DEFAULT, ParagraphAlignment.CENTER,"1212121212");
        builder.createTitle(1,"一、测试测试");
        builder.createTitle(2,"1.1 测试测试");
        builder.createTitle(3,"1.1.1 测试测试");
        builder.createTitle(4,"1.1.1.1 测试测试");
        builder.build(new FileOutputStream("/Users/jiangzhaoyue/Downloads/poiTest/1.docx"));
    }


}
