package com.xuxiong.easyexceldemo.easyexcel;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.*;

/**
 * @author 许兄
 * Created on 2019/12/11
 */
public class Test {

    public static void main(String[] args) {
        //////////////
        List<List<String>> dataStr = new ArrayList<>();
        List<String> strList = new ArrayList<>();
        strList.add("aaaa");
        dataStr.add(strList);

        List<Set<String>> includeColumnFiledNamesList = new ArrayList<>();
        Set<String> includeColumnFiledNames = new HashSet<String>(){{add("name");add("phone");}};
        includeColumnFiledNamesList.add(includeColumnFiledNames);
        Set<String> includeColumnFiledNames1 = new HashSet<String>(){{add("name");add("phone");add("email");add("address");}};
        includeColumnFiledNamesList.add(includeColumnFiledNames1);


        EasyExcelBuilder write = new EasyExcelBuilder()
                .buildDataModels(dataStr)
                .buildFileName("tttttx")
                .buildSheetName(Arrays.asList("aaa","bbb"))
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
