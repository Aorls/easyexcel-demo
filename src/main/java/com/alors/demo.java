package com.alors;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import lombok.Data;
import org.junit.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class demo {
    @Data
    @ColumnWidth(20)
    public static class TestVO {
        @ExcelProperty( value = "姓名",index = 0)
        private String name;
        @ExcelProperty( value = "年龄",index = 1)
        private int age;
        @ExcelProperty( value = "学校",index = 2)
        private String school;
    }
    /**
     * 多个sheet导入测试
     * @throws FileNotFoundException
     */
    @Test
    public void sheetImport() throws FileNotFoundException {
        // 标题
//        List<String> headList = Arrays.asList("姓名", "年龄", "学校");

        // 方法3 如果写到不同的sheet 不同的对象
//        String fileName = TestFileUtil.getPath() + "repeatedWrite" + System.currentTimeMillis() + ".xlsx";
        String fileName = "D:/2.xlsx";
        // 这里 指定文件
        ExcelWriter excelWriter = EasyExcel.write(fileName).build();
        WriteSheet writeSheet = null;
        // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来。这里最终会写到5个sheet里面
        for (int i = 0; i < 5; i++) {
            // 每次都要创建writeSheet 这里注意必须指定sheetNo。这里注意DemoData.class 可以每次都变，我这里为了方便 所以用的同一个class 实际上可以一直变
            writeSheet = EasyExcel.writerSheet(i, "模板"+i).head(TestVO.class).build();
            // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
            List<TestVO> data = data();
            excelWriter.write(data, writeSheet);
        }
        // 千万别忘记finish 会帮忙关闭流
        excelWriter.finish();
    }

    //模拟从数据库拿数据
    private List<TestVO> data() {
        List<TestVO> dataList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            TestVO testVO = new TestVO();
            testVO.setAge(i + 20);
            testVO.setName("vo" + i);
            testVO.setSchool("school" + i);
            dataList.add(testVO);
        }
        return dataList;
    }
}
