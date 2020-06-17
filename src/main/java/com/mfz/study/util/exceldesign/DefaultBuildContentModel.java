package com.mfz.study.util.exceldesign;

import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.springframework.util.Assert;
import org.springframework.util.CollectionUtils;

import java.util.Collection;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @Description: 内容设置-默认-处理Model 类型数据源
 * @Author mengfanzhu
 * @Date 2019/12/18 11:14
 * @Version 1.0
 */
@Slf4j
@NoArgsConstructor
public class DefaultBuildContentModel extends AbstractBuildContent {

    private int rowStartIndex = 0;

    public DefaultBuildContentModel(Map<String, CellStyle> styleMap, SXSSFSheet sheet, AbstractGetFieldValue abstractGetFieldValue) {
        super(styleMap, sheet, abstractGetFieldValue);
    }

    public DefaultBuildContentModel initParams(int rowStartIndex) {
        this.rowStartIndex = rowStartIndex;
        return this;
    }

    @Override
    void setContent(Collection data, String[] columns, LinkedHashMap<String, Integer> columnTypes) {
        modelHandler(data, columns);
    }

    private void modelHandler(Collection data, String[] columns) {
        Assert.notEmpty(columns, "columns cannot be null");
        if (CollectionUtils.isEmpty(data)) {
            return;
        }
        log.info("modelHandler start");
        SXSSFRow row = null;
        SXSSFCell cell = null;
        Object fieldValue = null;
        XSSFRichTextString xssfVal = null;
        Iterator iter = data.iterator();

        while (iter.hasNext()) {
            rowStartIndex++;

            row = sheet.createRow(rowStartIndex);
            Object t = iter.next();
            for (int i = 0, j = columns.length; i < j; i++) {
                cell = row.createCell(i);
                cell.setCellType(CellType.STRING);
                cell.setCellStyle(styleMap.get(AbstractBuildStyle.CONTENT_STYLE_KEY));

                fieldValue = abstractGetFieldValue.getFieldValue(t, columns[i]);
                xssfVal = new XSSFRichTextString(String.valueOf(fieldValue));
                cell.setCellValue(xssfVal);
            }
        }
    }
}
