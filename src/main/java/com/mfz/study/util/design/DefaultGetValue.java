package com.mfz.study.util.design;

/**
 * @Description: 默认获取属性值
 * @Author mengfanzhu
 * @Date 2019/12/18 10:36
 * @Version 1.0
 */
public class DefaultGetValue extends AbstractGetFieldValue{

    @Override
    Object getFieldValue(Object obj, String fieldName) {
        if(null == fieldName){
            return "";
        }
        return ReflectingHelper.getValueByField(obj,fieldName.toString());
    }
}
