package com.mfz.study.util.design;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.awt.Color;
import java.util.Map;

/**
 * @Description: 设置样式-默认
 * @Author mengfanzhu
 * @Date 2019/12/18 11:05
 * @Version 1.0
 */
public class DefaultBuildStyle extends AbstractBuildStyle {

    private static Color blue = new Color(135, 206, 255);

    @Override
    void setStyle(SXSSFWorkbook workbook, Map<String, CellStyle> styleMap) {
        XSSFCellStyle style1 = (XSSFCellStyle) workbook.createCellStyle();
        /**
         * 标题样式
         */
        // 前景色(蓝色)
        XSSFColor myColor0= new XSSFColor(blue);
        style1.setFillForegroundColor(myColor0);
        // 设置单元格填充样式
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 设置单元格边框
        style1.setBorderBottom(BorderStyle.THIN);
        style1.setBorderLeft(BorderStyle.THIN);
        style1.setBorderRight(BorderStyle.THIN);
        style1.setBorderTop(BorderStyle.THIN);
        style1.setAlignment(HorizontalAlignment.CENTER);
        // 生成一个字体
        XSSFFont font0 = (XSSFFont) workbook.createFont();
        font0.setFontHeightInPoints((short) 14);
        font0.setBold(true);
        // 把字体应用到当前的样式
        style1.setFont(font0);
        styleMap.put(TITLE_STYLE_KEY, style1);

        /**
         * 表头样式
         */
        // 生成一个字体
        XSSFFont font = (XSSFFont) workbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        // 把字体应用到当前的样式
        style1.setFont(font);
        styleMap.put(HEAD_STYLE_KEY, style1);

        /**
         * 内容样式
         */
        XSSFCellStyle style2 = (XSSFCellStyle) workbook.createCellStyle();
        XSSFColor white = new XSSFColor(Color.WHITE);
        style2.setFillBackgroundColor(white);
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBorderLeft(BorderStyle.THIN);
        style2.setBorderRight(BorderStyle.THIN);
        style2.setBorderTop(BorderStyle.THIN);
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        // 生成另一个字体
        XSSFFont font2 = (XSSFFont) workbook.createFont();
        font2.setBold(false);
        // 把字体应用到当前的样式
        style2.setFont(font2);
        styleMap.put(CONTENT_STYLE_KEY, style2);
    }
}
