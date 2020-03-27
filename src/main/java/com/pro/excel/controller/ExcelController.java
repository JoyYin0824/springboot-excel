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


    @RequestMapping("/uploadDutyExcelNew")
    @ResponseBody
    public String uploadDutyExcel(HttpServletRequest request,
                                  @RequestParam("file") MultipartFile multfile) throws Exception {
        // 获取文件名
        String fileName = multfile.getOriginalFilename();
        // 获取文件后缀
        String prefix=fileName.substring(fileName.lastIndexOf("."));
        // 用uuid作为文件名，防止生成的临时文件重复
        final File excelFile = File.createTempFile("", prefix);
        // MultipartFile to File
        multfile.transferTo(excelFile);

        //你的业务逻辑

        //程序结束时，删除临时文件
        deleteFile(excelFile);
        return "";
    }

    /**
     * 删除
     *
     * @param files
     */
    private void deleteFile(File... files) {
        for (File file : files) {
            if (file.exists()) {
                file.delete();
            }
        }
    }


}
