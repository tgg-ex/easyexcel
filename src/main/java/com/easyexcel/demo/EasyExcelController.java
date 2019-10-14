package com.easyexcel.demo;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author zyz
 * <p>
 * easyExcel demo控制器
 */
@RestController
public class EasyExcelController {


    @RequestMapping("/readDemo")
    public String readExcel() throws Exception {
        ExcelReader excelReader = new ExcelReader(new FileInputStream("D:\\test.xlsx"), ExcelTypeEnum.XLSX, null, new DemoDataListener());
        excelReader.read(new Sheet(1, 1, ReadData.class));
        return "ok";
    }

    @RequestMapping("/writeDemo")
    public String writeExcel() throws FileNotFoundException {
        OutputStream out = new FileOutputStream("D:\\test.xlsx");
        ExcelWriter writer = EasyExcelFactory.getWriter(out);

        // 写仅有一个 Sheet 的 Excel 文件, 此场景较为通用
        Sheet sheet1 = new Sheet(1, 0, ReadData.class);

        // 第一个 sheet 名称
        sheet1.setSheetName("第一个sheet");

        // 写数据到 Writer 上下文中
        // 入参1: 数据库查询的数据list集合
        // 入参2: 要写入的目标 sheet
        writer.write(getAllUser(), sheet1);

        // 将上下文中的最终 outputStream 写入到指定文件中
        writer.finish();
        return "ok";
    }

    public List<ReadData> getAllUser() {
        List<ReadData> userList = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            ReadData user = ReadData.builder().name("张三" + i).password("1234").age(i).build();
            userList.add(user);
        }
        return userList;
    }
}
