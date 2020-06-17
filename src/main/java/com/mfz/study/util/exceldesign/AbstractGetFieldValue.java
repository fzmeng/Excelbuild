package com.mfz.study.util.exceldesign;

/**
 * @Description: 获得属性值抽象
 * @Author mengfanzhu
 * @Date 2019/12/18 10:35
 * @Version 1.0
 */
public abstract class AbstractGetFieldValue {

    /**
     * 获取属性值方法
     * @param obj 反射对象
     * @param fieldName 属性名
     * @return
     */
    abstract Object getFieldValue(Object obj,String fieldName);
}
