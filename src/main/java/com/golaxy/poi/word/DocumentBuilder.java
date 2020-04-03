package com.golaxy.poi.word;

import com.golaxy.poi.word.bean.TitleStyle;
import com.golaxy.poi.word.bean.TitleStyleList;
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
    private TitleStyleList titleStyles;

    /**
     * 新建文档
     */
    public DocumentBuilder() {
        this.xwpfDocument = new XWPFDocument();
        this.titleStyles = new TitleStyleList();
    }

    /**
     * 根究word模板创建
     *
     * @param template
     * @throws IOException
     */
    public DocumentBuilder(InputStream template) throws IOException {
        this.xwpfDocument = new XWPFDocument(template);
    }

    /**
     * 根据自定义样式层级创建文档
     *
     * @param list
     */
    public DocumentBuilder(List<TitleStyle> list) {
        this.titleStyles = new TitleStyleList(list);
    }

    /**
     * 根究word模板创建
     *
     * @param template 模板输入流
     * @param list     title列表
     * @throws IOException
     */
    public DocumentBuilder(InputStream template, List<TitleStyle> list) throws IOException {
        this.xwpfDocument = new XWPFDocument(template);
        this.titleStyles = new TitleStyleList(list);
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

    /**
     * 按层级创建标题
     * @param level 层级，默认支持:1/2/3/4,字体大小默认为:18/16/14/12
     * @param text 文本内容
     */
    public void createTitle(int level, String text) {
        initTitle();
        XWPFParagraph paragraph = xwpfDocument.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        paragraph.setStyle(titleStyles.getNameByLevel(level));

    }

    public void build(OutputStream stream) throws IOException {
        xwpfDocument.write(stream);
    }

    /**
     * 加载标题样式
     */
    private void initTitle() {
        if (isInitTitle) {
            return;
        }
        titleStyles.getList().forEach(item -> {
            addCustomHeadingStyle(item.getName(), item.getLevel(), item.getFontSize());
        });
        isInitTitle = true;

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
