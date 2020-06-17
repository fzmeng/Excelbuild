package com.mfz.study.util.exceldesign;

import com.alibaba.fastjson.JSONObject;
import com.mfz.study.util.InvokeDataSource;
import lombok.Data;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

/**
 * @Description:
 * @Author mengfanzhu
 * @Date 2019/12/18 14:19
 * @Version 1.0
 */
@RestController
public class TestController {

    static List<ExcelExportModel> excelExportModels = new ArrayList<>();
    static List<LinkedHashMap<String, Object>> excelExportMaps = new ArrayList<>();

    static LinkedHashMap<String, Integer> columnTypeMap = new LinkedHashMap<String, Integer>() {{
        put("id", ExcelCellType.STRING.getCode());
        put("name", ExcelCellType.STRING.getCode());
        put("address", ExcelCellType.STRING.getCode());
    }};


    static LinkedHashMap<String, String> titleMap = new LinkedHashMap<String, String>() {{
        put("id", "标识");
        put("name", "名字");
        put("address", "地址");
    }};


    static {
        int max = 10000;
        for (int i = 0; i < max; i++) {
            ExcelExportModel excelExportModel = new ExcelExportModel();
            excelExportModel.setName("黎明" + i);
            excelExportModel.setAddress("地址" + i);
            excelExportModel.setId(i + "");
            excelExportModels.add(excelExportModel);
        }

        for (int i = 0; i < max; i++) {
            LinkedHashMap<String, Object> d = new LinkedHashMap<>();
            d.put("name", "黎明" + i);
            d.put("address", "北京朝阳" + i);
            d.put("id", "00001" + i);
            excelExportMaps.add(d);
        }
    }

    @Data
    public static class ExcelExportModel {

        private String id;

        private String name;

        private String address;
    }

    @GetMapping(value = "/export")
    public void export(HttpServletResponse response) {
        useWay1(response);
    }

    @GetMapping(value = "/export3")
    public void export3(HttpServletResponse response) {
        useWay3(response);
    }

    @GetMapping(value = "/export4")
    public void export4(HttpServletResponse response) {
        useWay4(response);
    }


    /**
     * invoke data
     *
     * @param response
     */
    private void useWay4(HttpServletResponse response) {
        ExcelExportModel param = new ExcelExportModel();
        param.setId("10");
        param.setName("我是反射获取数据源");
        param.setAddress("反射类地址");

        new ExcelDirector()
                .addFieldTitleTable(titleMap)
                .setExcelFileName("反射获取数据源文件.xlsx")
                .setColumnTypeMap(columnTypeMap)
                .setBoundInvokeClass(InvokeDataSource.class)
                .setBoundInvokeMethodParamValues(JSONObject.toJSONString(param), 0, 1000)
                .constructMapInvokeData().flush(response);
    }

    /**
     * Map 导出
     *
     * @param response
     */
    private void useWay3(HttpServletResponse response) {
        new ExcelDirector()
                .addFieldTitleTable(titleMap)
                .setData(excelExportMaps)
                .setExcelFileName("c201900202")
                //可不指定columnTypeMap ，默认为String
//                .setColumnTypeMap(columnTypeMap)
                .constructMap()
                .flush(response);
    }


    /**
     * model 1
     *
     * @param response
     */
    private void useWay1(HttpServletResponse response) {
        //实体对象导出示例
        new ExcelDirector()
                .addFieldTitleTable(titleMap)
                .setData(excelExportModels)
                .setExcelFileName("测试Excel.xlsx")
                .setTitles(new String[]{"北京统计", "20191229"})
                .constructModel()
                .flush(response);
    }
}
