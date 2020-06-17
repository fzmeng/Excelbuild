package com.mfz.study.util.exceldesign;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.util.Assert;
import org.springframework.util.StringUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.Collection;
import java.util.LinkedHashMap;

/**
 * @Description: 建造者模式-具体建造者
 * @Author mengfanzhu
 * @Date 2019/12/18 11:39
 * @Version 1.0
 */
@Slf4j
public class ExcelBuilder extends AbstractExcelBuilder {

    /**
     * 构建滑动窗口数
     * new SXSSFWorkbook() -> 默认为100
     */
    private static final Integer ROW_ACCESS_WINDOWSIZE = 200;

    @Override
    AbstractExcelBuilder buildWorkbook(SXSSFWorkbook wb) {
        if (null == workbook) {
            workbook = new SXSSFWorkbook(ROW_ACCESS_WINDOWSIZE);
        } else {
            workbook = wb;
        }
        return this;
    }

    @Override
    AbstractExcelBuilder buildSheet(String sheetName) {
        if (StringUtils.isEmpty(sheetName)) {
            sheetName = "sheet1";
        }
        sheet = workbook.createSheet(sheetName);
        return this;
    }

    @Override
    AbstractExcelBuilder buildContent(Collection data, String[] heads, LinkedHashMap<String, Integer> columnTypeMap, Class<? extends AbstractBuildContent> abstractSetContent) {
        setContentImpl(abstractSetContent);
        this.content.setContent(data, heads, columnTypeMap);
        return this;
    }

    @Override
    AbstractExcelBuilder buildTitlesAndHeads(String[] titles, String[] heads) {
        Assert.notEmpty(heads, "buildTitlesAndHeads->heads cannot be null");
        if (null != titles && titles.length > 0) {
            super.rowStartIndex = 1;
        }
        abstractBuildTitle.setTitle(sheet, titles, styleMap, heads, fieldTitleTable);
        return this;
    }

    @Override
    AbstractExcelBuilder buildExcelFileName(String excelFileName) {
        Assert.notNull(excelFileName, "excelFileName cannot be null");
        return this;
    }

    @Override
    AbstractExcelBuilder buildStyle() {
        abstractBuildStyle.setStyle(workbook, styleMap);
        return this;
    }

    @Override
    AbstractExcelBuilder buildInvokeClass(Class clazz) {
        super.invokeClass = clazz;
        return this;
    }

    @Override
    AbstractExcelBuilder buildInvokeMethodParamValue(String methodParamJson, Integer methodParamLimitStart, Integer methodParamLimitSize) {
        super.methodParamJson = methodParamJson;
        super.methodParamLimitStart = methodParamLimitStart;
        super.methodParamLimitSize = methodParamLimitSize;
        return this;
    }


    /**
     * 设定每次从内存读多少
     */
    private static byte[] buff = new byte[2048];

    @Override
    void flush(HttpServletResponse response) {
        flush(response, this.excelFileName);
    }

    @Override
    void flush(HttpServletResponse response, String excelFileName) {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try {
            workbook.write(os);
        } catch (Exception e) {
            e.printStackTrace();
            log.error(e.toString());
        }
        //设置response参数，可以打开下载页面
        packResponse(response, excelFileName);
        try (
                // 封装为缓冲，减少读写次数，提高效率
                BufferedInputStream bis = new BufferedInputStream(new ByteArrayInputStream(os.toByteArray()));
                BufferedOutputStream bos = new BufferedOutputStream(response.getOutputStream());
        ) {
            //循环从内存读取，并写到Resonse响应流中
            int bytesRead;
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }
        } catch (IOException e) {
            log.error(e.toString());
        } finally {
            // 将此workbook对应的临时文件删除
            if (null != workbook) {
                workbook.dispose();
            }
        }
    }

    private void packResponse(HttpServletResponse response, String excelFileName) {
        try {
            if(StringUtils.isEmpty(excelFileName)){
                excelFileName = "excel-"+System.currentTimeMillis()+".xlsx";
            }
            response.reset();
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + new String(excelFileName.getBytes(), "iso-8859-1"));
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
    }
}

