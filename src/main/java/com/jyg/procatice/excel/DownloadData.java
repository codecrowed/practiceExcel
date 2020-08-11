package com.jyg.procatice.excel;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.util.Date;

@Data
public class DownloadData {
    @ExcelProperty(value = {"信息","字符串标题"})
    private String string;
    @ExcelProperty(value = {"信息","日期标题"})
    private Date date;
    @ExcelProperty(value = {"信息","数字标题"})
    private Double doubleData;
}
