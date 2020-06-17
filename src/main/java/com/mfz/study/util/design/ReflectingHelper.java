package com.mfz.study.util.design;


import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.formula.functions.T;
import org.springframework.util.StringUtils;

import java.lang.reflect.Field;

/**
 * @Description: 反射辅助类
 * @Author mengfanzhu
 * @Date 2019/12/18 10:38
 * @Version 1.0
 */
@Slf4j
public abstract class ReflectingHelper {

    public static Object getValueByField(Object obj, String fieldName) {
        if (obj == null || StringUtils.isEmpty(fieldName)) {
            return null;
        }

        Field field = getFieldByFieldName(obj, fieldName);
        Object val = null;

        if (!field.isAccessible()) {
            field.setAccessible(true);
        }
        try {
            val = field.get(obj);
            return val;
        } catch (IllegalAccessException e) {
            e.printStackTrace();
            log.error("反射辅助类-获取值异常 " + e.toString());
        } finally {
            field.setAccessible(false);
        }
        return null;
    }

    /**
     * @param obj
     * @param fieldName
     * @return
     */
    public static Field getFieldByFieldName(Object obj, String fieldName) {
        for (Class<?> superClass = obj.getClass(); superClass != T.class; superClass = superClass.getSuperclass()) {
            try {
                return superClass.getDeclaredField(fieldName);
            } catch (NoSuchFieldException e) {
                e.printStackTrace();
                log.error("反射辅助类-getFieldByFieldName-Exception:" + e.toString());
            }
        }
        return null;
    }
}
