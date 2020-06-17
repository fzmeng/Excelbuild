package com.mfz.study.util.exceldesign;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @Description: 设置内容-抽象
 * @Author mengfanzhu
 * @Date 2019/12/18 11:11
 * @Version 1.0
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public abstract class AbstractBuildContent {

    Map<String, CellStyle> styleMap;

    SXSSFSheet sheet;

    AbstractGetFieldValue abstractGetFieldValue;

    /**
     * @param data
     * @param columns
     * @param columnTypes
     */
    abstract void setContent(Collection data, String[] columns, LinkedHashMap<String, Integer> columnTypes);
}
