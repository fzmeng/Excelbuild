package com.mfz.study.util.design;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.springframework.util.Assert;

import java.util.Map;

/**
 * @Description: 默认-设置标题-表头
 * @Author mengfanzhu
 * @Date 2019/12/18 10:57
 * @Version 1.0
 */
public class DefaultBuildTitle extends AbstractBuildTitle {

    @Override
    void setTitle(SXSSFSheet sheet, String[] titles, Map<String, CellStyle> styleMap, String[] heads, Map<String, String> fieldTitleTable) {
        int rowIndex = 0;
        if (null != titles && titles.length != 0) {
            SXSSFRow rowTitle = sheet.createRow(rowIndex);

            for (int i = 0, j = titles.length; i < j; i++) {
                SXSSFCell cell = rowTitle.createCell(i);
                cell.setCellType(CellType.STRING);
                cell.setCellStyle(styleMap.get(AbstractBuildStyle.TITLE_STYLE_KEY));
                cell.setCellValue(titles[i]);
            }

            rowIndex++;
        }

        if (null != heads && heads.length != 0) {
            Assert.notEmpty(fieldTitleTable, "fieldTitleTable不能为空！");

            SXSSFRow rowHead = sheet.createRow(rowIndex);
            for (int i = 0, j = heads.length; i < j; i++) {
                SXSSFCell cell = rowHead.createCell(i);
                cell.setCellType(CellType.STRING);
                cell.setCellStyle(styleMap.get(AbstractBuildStyle.HEAD_STYLE_KEY));

                XSSFRichTextString textString = new XSSFRichTextString(fieldTitleTable.get(heads[i]));
                cell.setCellValue(textString);
            }
        }
    }
}
