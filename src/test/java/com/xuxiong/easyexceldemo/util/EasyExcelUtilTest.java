package com.xuxiong.easyexceldemo.util;

import com.xuxiong.easyexceldemo.util.easyexcel.Demo;
import com.xuxiong.easyexceldemo.util.easyexcel.EasyExcelBuilder;
import com.xuxiong.easyexceldemo.util.easyexcel.EasyExcelUtil;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.jupiter.api.Test;

import java.util.*;

/**
 * @author 许兄
 * Created on 2019/12/12
 */
public class EasyExcelUtilTest {

    @Test
    public void writeExcel(){
        //设置需要导出的字段名
        List<Set<String>> includeColumnFiledNamesList = new ArrayList<>();
        Set<String> includeColumnFiledNames = new HashSet<String>(){{add("name");add("phone");}};
        includeColumnFiledNamesList.add(includeColumnFiledNames);
        Set<String> includeColumnFiledNames1 = new HashSet<String>(){{add("name");add("phone");add("email");add("address");}};
        includeColumnFiledNamesList.add(includeColumnFiledNames1);


        //根据需要设置模式
        EasyExcelBuilder write = new EasyExcelBuilder()
                .buildFileName("文件名称")
                .buildSheetName(Arrays.asList("第一个sheet","第二个sheet"))
//                .buildOssPath("excel/default/201906/")
                .buildIncludeColumnFiledNames(includeColumnFiledNamesList)
//                .buildExcludeColumnFiledNames(includeColumnFiledNamesList)
                .buildFilePath("/Users/xuxiong/Downloads/")
                .buildHeadColor(IndexedColors.LIGHT_YELLOW)
//                .buildContentColor(IndexedColors.RED)
                .build();
        List<List<?>> data = new ArrayList<>();
        data.add(data("a列"));
        data.add(data("b列"));
        EasyExcelUtil.writeExcel(write, Demo.class, data);
    }
    private static List<Demo> data(String text){
        List<Demo> demos = new ArrayList<>();
        Demo demo;
        for(int i = 0; i < 10; i++){
            demo = new Demo();
            demo.setName("name" + i + "_"+text);
            demo.setAge(i);
            demo.setAddress("address" + i+ "_"+text);
            demo.setEmail("email" + i+ "_"+text);
            demo.setPhone("phone" + i+ "_"+text);
            demos.add(demo);
        }
        return demos;
    }
}
