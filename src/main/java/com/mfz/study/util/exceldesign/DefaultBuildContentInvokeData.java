package com.mfz.study.util.exceldesign;

import com.mfz.study.util.SpringContextHolder;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.springframework.context.ApplicationContext;

import java.lang.reflect.Method;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @Description: 内容设置-默认-处理Model 类型数据源
 * 适用于分页循环获取数据源-输入到文件内容中
 * @Author mengfanzhu
 * @Date 2019/12/18 11:14
 * @Version 1.0
 */
@Slf4j
@NoArgsConstructor
public class DefaultBuildContentInvokeData extends AbstractBuildContent {

    private int rowStartIndex = 0;
    private Class invokeClass;

    /**
     * @see BaseSetContentInterface
     */
    private String json;
    private Integer limitStart;
    private Integer limitSize;

    public DefaultBuildContentInvokeData(Map<String, CellStyle> styleMap, SXSSFSheet sheet, AbstractGetFieldValue abstractGetFieldValue) {
        super(styleMap, sheet, abstractGetFieldValue);
    }

    public DefaultBuildContentInvokeData initParams(Class invokeClass, String methodParamJson, Integer methodParamLimitStart, Integer methodParamLimitSize, int rowStartIndex) {
        this.rowStartIndex = rowStartIndex;
        this.invokeClass = invokeClass;
        this.json = methodParamJson;
        this.limitStart = methodParamLimitStart;
        this.limitSize = methodParamLimitSize;
        return this;
    }

    @Override
    void setContent(Collection data, String[] columns, LinkedHashMap<String, Integer> columnTypes) {
        try {
            if (!(invokeClass.newInstance() instanceof BaseSetContentInterface)) {
                throw new Exception("DefaultBuildContentInvokeData.invokeClass need to implement BaseSetContentInterface ！");
            }
            log.info("DefaultBuildContentInvokeData setContent start,invokeClasss is {},json is {},limitStart is {},limitSize is {} .", invokeClass, json, limitStart, limitSize);

            Method invokeMethod = BaseSetContentInterface.class.getMethods()[0];

            ApplicationContext context = SpringContextHolder.getApplicationContext();
            Object bean = context.getBean(invokeClass);
            /**
             * 给实例化的类注入需要的bean (@Autowired) ,如果不注入，被@Autowired注解的变量会报空指针
             */
            context.getAutowireCapableBeanFactory().autowireBean(bean);
            data = (Collection) invokeMethod.invoke(bean, json, limitStart, limitSize);

            DefaultBuildContentMap defaultSetContentMap = new DefaultBuildContentMap(styleMap, sheet, abstractGetFieldValue).initParams(rowStartIndex);
            for (; data.size() > 0;
                 limitStart += limitSize, data = (Collection) invokeMethod.invoke(bean, json, limitStart, limitSize)) {
                defaultSetContentMap.mapHandler(data, columnTypes);
            }
        } catch (Exception e) {
            log.error("DefaultBuildContentInvokeData exception", e.getMessage());
        }
    }
}
