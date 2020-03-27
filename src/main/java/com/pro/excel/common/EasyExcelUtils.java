package com.pro.excel.common;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import lombok.Data;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * @author :Mr.kk
 */
public class EasyExcelUtils {
    private static final Logger logger = LoggerFactory.getLogger(EasyExcelUtils.class);

    private static Sheet initSheet;

    static {
        initSheet = new Sheet(1, 0);
        initSheet.setSheetName("sheet");
        // 设置自适应宽度
        initSheet.setAutoWidth(Boolean.TRUE);
    }

    /**
     * 读取少于1000行数据
     *
     * @param is 文件流
     * @return
     */
    public static List<Object> readLessThan1000Row(InputStream is) {
        return readLessThan1000RowBySheet(is, null);
    }

    /**
     * 读取少于1000行数据，带样式的
     *
     * @param is 文件流
     * @param sheet
     * @return
     */
    public static List<Object> readLessThan1000RowBySheet(InputStream is, Sheet sheet) {
        sheet = sheet != null ? sheet : initSheet;
        try {
            return EasyExcelFactory.read(is, sheet);
        } catch (Exception e) {
            logger.error("发生异常", e);
        } finally {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (IOException e) {
                logger.error("excel文件读取失败，失败原因：{}", e);
            }
        }
        return null;
    }

    /**
     * 读取大于1000行数据
     * @param filePath
     * @return
     */
    public static List<Object> readMoreThan1000Row(String filePath) {
        return readMoreThan1000RowBySheet(filePath, null);
    }

    /**
     * 读取大于1000行数据
     * @param filePath
     * @param sheet
     * @return
     */
    public static List<Object> readMoreThan1000RowBySheet(String filePath, Sheet sheet) {
        if (!StringUtils.hasText(filePath)) {
            return null;
        }
        sheet = sheet != null ? sheet : initSheet;
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(filePath);
            ExcelListener excelListener = new ExcelListener();
            EasyExcelFactory.readBySax(inputStream, sheet, excelListener);
            return excelListener.getDatas();
        } catch (FileNotFoundException e) {
            logger.error("找不到文件或者文件路径错误");
        } finally {
            try {
                if (inputStream != null) {
                    inputStream.close();
                }
            } catch (IOException e) {
                logger.error("excel文件读取失败，失败原因：{}", e);
            }
        }
        return null;
    }

    /**
     * 导出单个sheet
     * @param response
     * @param dataList
     * @param sheet
     * @param fileName
     * @throws UnsupportedEncodingException
     */
    public static void writeExcelOneSheet(HttpServletResponse response, List<? extends BaseRowModel> dataList, Sheet sheet, String fileName) throws UnsupportedEncodingException {
        if (CollectionUtils.isEmpty(dataList)) {
            return;
        }
        // 如果sheet为空，则使用默认的
        if (null == sheet) {
            sheet = initSheet;
        }
        try {
            String value = "attachment; filename=" + new String(
                    (fileName + ExcelTypeEnum.XLSX.getValue()).getBytes("gb2312"), "ISO8859-1");
            response.setContentType("multipart/form-data");
            response.setCharacterEncoding("utf-8");
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            response.setHeader("Content-disposition", value);

            ServletOutputStream out = response.getOutputStream();
            ExcelWriter writer = EasyExcelFactory.getWriter(out, ExcelTypeEnum.XLS, true);
            // 设置属性类
            sheet.setClazz(dataList.get(0).getClass());
            writer.write(dataList, sheet);
            writer.finish();
            out.flush();
        } catch (IOException e) {
            logger.error("导出失败，失败原因：{}", e);
        }
    }

    /**
     * @Author lockie
     * @Description 导出excel 支持一张表导出多个sheet
     * @Param OutputStream 输出流
     * Map<String, List>  sheetName和每个sheet的数据
     * ExcelTypeEnum 要导出的excel的类型 有ExcelTypeEnum.xls 和有ExcelTypeEnum.xlsx
     * @Date 上午12:16 2019/1/31
     */
    public static void writeExcelMutilSheet(HttpServletResponse response, Map<String, List<? extends BaseRowModel>> dataList, String fileName) throws UnsupportedEncodingException {
        if (CollectionUtils.isEmpty(dataList)) {
            return;
        }
        try {
            String value = "attachment; filename=" + new String(
                    (fileName + new SimpleDateFormat("yyyyMMdd").format(new Date()) + ExcelTypeEnum.XLSX.getValue()).getBytes("gb2312"), "ISO8859-1");
            response.setContentType("multipart/form-data");
            response.setCharacterEncoding("utf-8");
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            response.setHeader("Content-disposition", value);
            ServletOutputStream out = response.getOutputStream();
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX, true);
            // 设置多个sheet
            setMutilSheet(dataList, writer);
            writer.finish();
            out.flush();
        } catch (IOException e) {
            logger.error("导出异常", e);
        }
    }


    /**
     * @Author lockie
     * @Description //setSheet数据
     * @Date 上午12:39 2019/1/31
     */
    private static void setMutilSheet(Map<String, List<? extends BaseRowModel>> dataList, ExcelWriter writer) {
        int sheetNum = 1;
        for (Map.Entry<String, List<? extends BaseRowModel>> stringListEntry : dataList.entrySet()) {
            Sheet sheet = new Sheet(sheetNum, 0, stringListEntry.getValue().get(0).getClass());
            sheet.setSheetName(stringListEntry.getKey());
            writer.write(stringListEntry.getValue(), sheet);
            sheetNum++;
        }
    }



    /**
     * 导出监听
     */
    @Data
    public static class ExcelListener extends AnalysisEventListener {
        private List<Object> datas = new ArrayList<>();

        /**
         * 逐行解析
         * @param object 当前行的数据
         * @param analysisContext
         */
        @Override
        public void invoke(Object object, AnalysisContext analysisContext) {
            // 当前行
//            analysisContext.getCurrentRowNum()
            if (object != null) {
                datas.add(object);
            }
        }


        /**
         * 解析完所有数据后会调用该方法
         * @param analysisContext
         */
        @Override
        public void doAfterAllAnalysed(AnalysisContext analysisContext) {

        }
    }

    ///**
    // * test
    // */
    //public static void main(HttpServletResponse response) throws UnsupportedEncodingException {
    //    // 导出多个sheet
    //    //List<OnlineDataExportDTO> onlineDataExportDTOS = new ArrayList<OnlineDataExportDTO>();
    //    //OnlineDataExportDTO onlineDataExportDTO = new OnlineDataExportDTO("lslsls","1");
    //    //onlineDataExportDTOS.add(onlineDataExportDTO);
    //    //Map<String, List<? extends BaseRowModel>> map = new HashMap<>();
    //    //map.put("自营订单", onlineDataExportDTOS);
    //    //map.put("互联互通", onlineDataExportDTOS);
    //    //String fileName = new String(("测试导出2019").getBytes(), "UTF-8");
    //    //writeExcelMutilSheet(response, map, fileName);
    //
    //    // 导出单个sheet
    //    //writeExcelOneSheet(response, onlineDataExportDTOS, null, fileName);
    //}
}
