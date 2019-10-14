package com.easyexcel.demo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Builder;
import lombok.Data;

import java.util.Date;

/**
 * @author zyz
 */
@Data
@Builder
public class ReadData extends BaseRowModel {
    @ExcelProperty(value = "姓名", index = 0)
    private String name;

    @ExcelProperty(value = "密码", index = 1)
    private String password;

    @ExcelProperty(value = "年龄", index = 2)
    private Integer age;

}
