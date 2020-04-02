package com.golaxy.poi.word;

import com.golaxy.poi.word.bean.TitleStyle;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.List;

/**
 * 快速创建文档
 *
 * @author jiangzhaoyue
 */
public class DocumentBuilder {
    private XWPFDocument xwpfDocument;
    private boolean isInitTitle = false;

    /**
     * 新建文档
     */
    public DocumentBuilder() {
        xwpfDocument = new XWPFDocument();
    }

    /**
     * 根究word模板创建
     *
     * @param template
     * @throws IOException
     */
    public DocumentBuilder(InputStream template) throws IOException {
        xwpfDocument = new XWPFDocument(template);
    }

    /**
     * 设置word的页眉
     *
     * @param type  DEFAULT表示默认,EVEN表示每页重复，FIRST表示仅首页
     * @param align 设置对齐
     * @param text
     */
    public void setHeader(HeaderFooterType type, ParagraphAlignment align, String text) {
        XWPFParagraph paragraph = xwpfDocument.createHeader(HeaderFooterType.DEFAULT).createParagraph();
        paragraph.setFontAlignment(align.getValue());
        paragraph.createRun().setText(text);
    }

    /**
     * 创建文本段落
     *
     * @param isFirstLineIndentation 是否首行缩进
     * @param fontSize               字体大小默认五号,10.5
     * @param text                   文本
     */
    public void creatParagraph(boolean isFirstLineIndentation, Integer fontSize, String text) {
        XWPFParagraph paragraph = xwpfDocument.createParagraph();
        if (isFirstLineIndentation) {
            //缇(Twips) （缇：计量单位，等于“磅”的 1/20，英寸的 1/1,440。一厘米有 567 缇。
            paragraph.setIndentationFirstLine((int) (20 * (fontSize != null ? fontSize : 10.5) * 2));
        }
        XWPFRun run = paragraph.createRun();
        if (fontSize != null) {
            run.setFontSize(fontSize);
        }
        run.setText(text);

    }

    public void createTitle() {
        initTitle(null);
        XWPFParagraph paragraph = xwpfDocument.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("标题 1");
        paragraph.setStyle("标题 1");

        XWPFParagraph paragraph2 = xwpfDocument.createParagraph();
        XWPFRun run2 = paragraph2.createRun();
        run2.setText("标题 1");
        paragraph2.setStyle("标题 2");

        XWPFParagraph paragraph3 = xwpfDocument.createParagraph();
        XWPFRun run3 = paragraph3.createRun();
        run3.setText("正文");

    }

    public void build(OutputStream stream) throws IOException {
        xwpfDocument.write(stream);
    }

    /**
     * 加载标题样式
     * @param titleStyleList 自定义时可以传入，需在首次插入前加载。
     */
    public void initTitle(List<TitleStyle> titleStyleList) {
        if (isInitTitle) {
            return;
        }
        if (titleStyleList != null && titleStyleList.size() > 0) {
            titleStyleList.forEach(item -> {
                addCustomHeadingStyle(item.getName(), item.getLevel(), item.getFontSize());
            });
        } else {
            addCustomHeadingStyle("标题 1", 1, 22);
            addCustomHeadingStyle("标题 2", 2, 16);
            addCustomHeadingStyle("标题 3", 3, 14);
            addCustomHeadingStyle("标题 4", 4, 12);
        }
    }

    /**
     * 增加自定义标题样式。这里用的stackoverflow的源码改造而来
     *
     * @param strStyleId   样式名称
     * @param headingLevel 样式级别
     * @param fontSize     样式字体大小
     */
    private void addCustomHeadingStyle(String strStyleId, int headingLevel, Integer fontSize) {

        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);

        // style defines a heading of the given level
        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);
        // 设置字体大小
        if (fontSize != null) {
            CTRPr rpr = CTRPr.Factory.newInstance();
            BigInteger bint = new BigInteger(Integer.toString(fontSize));
            CTHpsMeasure ctSize = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
            ctSize.setVal(bint.multiply(new BigInteger("2")));
            ctStyle.setRPr(rpr);
        }
        XWPFStyle style = new XWPFStyle(ctStyle);

        // is a null op if already defined
        XWPFStyles styles = xwpfDocument.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);

    }
}
