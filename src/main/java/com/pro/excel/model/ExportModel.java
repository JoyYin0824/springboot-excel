package com.pro.excel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @author :Mr.kk
 * @date: 2020/3/27 9:26
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class ExportModel extends BaseRowModel {

    @ExcelProperty(value = "编号", index = 0)
    private int num;

    @ExcelProperty(value = "姓名", index = 1)
    private String name;

    @ExcelProperty(value = "年龄", index = 2)
    private int age;






}
