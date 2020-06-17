package com.mfz.study.util.design;

import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.Map;

/**
 * @Description: 设置标题抽象
 * @Author mengfanzhu
 * @Date 2019/12/18 10:55
 * @Version 1.0
 */
@Getter
@Setter
public abstract class AbstractBuildTitle {

    /**
     * @param sheet           表格
     * @param titles          标题数组
     * @param heads           表头列数组
     * @param styleMap        样式库
     * @param fieldTitleTable 列对应表头名
     */
    abstract void setTitle(SXSSFSheet sheet, String[] titles, Map<String, CellStyle> styleMap, String[] heads, Map<String, String> fieldTitleTable);
}
