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
        builder.createTitle();
        builder.build(new FileOutputStream("/Users/jiangzhaoyue/Downloads/poiTest/1.docx"));
    }


}
