package com.mfz.study.util.exceldesign;

import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.springframework.util.Assert;
import org.springframework.util.CollectionUtils;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Description: 内容设置-默认-处理Map 类型数据源
 * @Author mengfanzhu
 * @Date 2019/12/18 11:14
 * @Version 1.0
 */
@Slf4j
@NoArgsConstructor
public class DefaultBuildContentMap extends AbstractBuildContent {

    private int rowStartIndex = 0;

    public DefaultBuildContentMap(Map<String, CellStyle> styleMap, SXSSFSheet sheet, AbstractGetFieldValue abstractGetFieldValue) {
        super(styleMap, sheet, abstractGetFieldValue);
    }

    public DefaultBuildContentMap initParams(int rowStartIndex){
        this.rowStartIndex = rowStartIndex;
        return this;
    }

    @Override
    void setContent(Collection data, String[] columns, LinkedHashMap<String, Integer> columnTypes) {
        mapHandler(data, columnTypes);
    }


    public void mapHandler(Collection data, LinkedHashMap<String, Integer> columnTypes) {
        Assert.notEmpty(columnTypes, "columnTypes cannot be null");
        if (CollectionUtils.isEmpty(data)) {
            return;
        }
        log.info("mapHandler start");
        SXSSFRow row = null;
        SXSSFCell cell = null;
        Iterator iter = data.iterator();
        int cellIndex = 0;
        Object fieldValue = null;
        Integer fieldType = null;
        //row loop
        while (iter.hasNext()) {
            Object t = iter.next();
            if (!(t instanceof Map)) {
                throw new IllegalArgumentException("DefaultBuildContentMap->data type must is List<Map<String,Object>> !");
            }
            rowStartIndex++;
            cellIndex = 0;
            row = sheet.createRow(rowStartIndex);
            Map<String, Object> map = (Map<String, Object>) t;
            //cell loop
            for (String fieldName : columnTypes.keySet()) {
                fieldValue = map.get(fieldName);
                if(null == fieldValue){
                    cellIndex++;
                    continue;
                }
                cell = row.createCell(cellIndex);
                cell.setCellType(CellType.STRING);
                cell.setCellStyle(styleMap.get(AbstractBuildStyle.CONTENT_STYLE_KEY));
                fieldType = columnTypes.get(fieldName);
                switch (ExcelCellType.getByCode(fieldType)) {
                    case STRING:
                        if (fieldValue instanceof String) {
                            String cellValue = String.valueOf(fieldValue);
                            cell.setCellValue(cellValue);
                        }
                        break;
                    case DECIMAL:
                        if (fieldValue instanceof BigDecimal) {
                            BigDecimal cellValue = (BigDecimal) fieldValue;
                            cellValue.setScale(2, BigDecimal.ROUND_DOWN);
                            cell.setCellValue(cellValue.toString());
                        }
                        break;
                    case _NONE:
                        break;
                    case TIME:
                        Timestamp resultTime = (Timestamp) fieldValue;
                        cell.setCellValue(getSimpleStringDate(resultTime));
                        break;
                    case DATE:
                        Date resultDate = (Date) fieldValue;
                        cell.setCellValue(getSimpleStringDate(resultDate));
                        break;
                }
                cellIndex++;
            }

        }
    }

    public static String getSimpleStringDate(Date date) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        sdf.setTimeZone(TimeZone.getTimeZone("Asia/Shanghai"));
        return sdf.format(date);
    }
}
