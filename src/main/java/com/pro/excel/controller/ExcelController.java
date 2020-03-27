package com.pro.excel.controller;

import com.pro.excel.common.EasyExcelUtils;
import com.pro.excel.model.ExportModel;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @author :Mr.kk
 * @date: 2020/3/27 9:19
 */
@RestController
@RequestMapping("/importAndExport")
public class ExcelController {


    /**
     * excel的导入
     */
    @GetMapping("/importExcel")
    public Object importExcel(@RequestParam(name = "file",required = true) MultipartFile excl){
        if(!excl.isEmpty()){//说明文件不为空
            try {
                InputStream is = new BufferedInputStream(excl.getInputStream());
                List<Object> list = EasyExcelUtils.readLessThan1000Row(is);
                //首先是读取行 也就是一行一行读，然后在取到列，遍历行里面的行，根据行得到列的值
                for (Object obj: list) {
                    System.out.println(obj);
                }
            }catch (Exception e){
                e.printStackTrace();
            }
        }
        return "";
    }


    /**
     * excel的导入
     */
    @GetMapping("/exportExcel")
    public void exportExcel(HttpServletRequest request, HttpServletResponse response){
        String fileName = null;
        try {
            List<ExportModel> dataList = new ArrayList<>();
            for (int i =0 ;i <= 3 ;i ++){
                dataList.add(new ExportModel(i,"Mr.kk"+i,18+i));
            }
            fileName = new String("excel导出".getBytes(), "UTF-8");
            EasyExcelUtils.writeExcelOneSheet(response,dataList,null,fileName);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
    }





}
