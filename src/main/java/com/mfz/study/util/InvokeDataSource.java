package com.mfz.study.util;

import com.alibaba.fastjson.JSONObject;
import com.mfz.study.util.design.BaseSetContentInterface;
import com.mfz.study.util.design.TestController;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Description: 反射数据源实现
 * @Author mengfanzhu
 * @Date 2019/12/18 14:19
 * @Version 1.0
 */
@Service
public class InvokeDataSource implements BaseSetContentInterface {

    @Override
    public List<? extends Map<String, Object>> getExcelExportData(String json, Integer limitStart, Integer size) {
        TestController.ExcelExportModel param = JSONObject.parseObject(json,new TestController.ExcelExportModel().getClass());
        System.out.println(param.getAddress());
        System.out.println(param.getId());
        System.out.println(param.getName());
        //TODO do something ...  根据入参 获取数据源

        if(limitStart >10){
            return new ArrayList<>();
        }
        List<HashMap<String, Object>> data = new ArrayList<>();
        for (int i = 0; i < size; i++) {
            HashMap<String, Object> d = new HashMap<>();
            d.put("name", "黎明" + i);
            d.put("address", "北京朝阳" + i);
            d.put("id", "00001" + i);
            data.add(d);
        }
        return data;
    }
}
