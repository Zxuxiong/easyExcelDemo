package com.xuxiong.easyexceldemo.util.easyexcel;


import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.Collections;
import java.util.List;
import java.util.Set;

/**
 * 风格建造者
 * @author 许兄
 * Created on 2019/12/11
 */
public class EasyExcelBuilder {

    /**
     * 文件名称 （最好加上后缀）
     */
    private String fileName = "模板.xlsx";
    /**
     * sheet
     */
    private List<String> sheetNames = Collections.singletonList("模板");
    /**
     * 数据类型对应的数据
     */
    private List<List<String>> dataModels;

    /**
     * 特定字段导出支持多个sheet 如果有这个数据即会忽略excludeColumnFiledNames
     */
    private List<Set<String>> includeColumnFiledNames;
    /**
     * 忽略特定字段导出 支持多个sheet、
     */
    private List<Set<String>> excludeColumnFiledNames;

    /**
     * 阿里oos存储路径（不包含文件名）无该路径时默认不上传到oos
     */
    private String ossPath;

    /** 服务器文件路径 */
    private String filePath = "/data/export/excel/";

    /** 表头颜色 目前只支持统一颜色 */
    private IndexedColors headColor = IndexedColors.GREY_25_PERCENT;

    /** 内容颜色 目前也只支持统一颜色 */
    private IndexedColors contentColor;


    public EasyExcelBuilder buildFileName(String fileName) {
        String newFileName = fileName.replaceAll("/", "");
        if (!newFileName.contains(".")) {
            newFileName += ".xlsx";
        }
        this.fileName = newFileName;
        return this;
    }

    public EasyExcelBuilder buildSheetName(List<String> sheetNames) {
        this.sheetNames = sheetNames;
        return this;
    }

    public EasyExcelBuilder buildDataModels(List<List<String>> dataModels) {
        this.dataModels = dataModels;
        return this;
    }

    public EasyExcelBuilder buildExcludeColumnFiledNames(List<Set<String>> excludeColumnFiledNames) {
        this.excludeColumnFiledNames = excludeColumnFiledNames;
        return this;
    }

    public EasyExcelBuilder buildIncludeColumnFiledNames(List<Set<String>> includeColumnFiledNames) {
        this.includeColumnFiledNames = includeColumnFiledNames;
        return this;
    }

    public EasyExcelBuilder buildOssPath(String ossPath) {
        this.ossPath = ossPath;
        return this;
    }

    public EasyExcelBuilder buildFilePath(String filePath) {
        this.filePath = filePath;
        return this;
    }

    public EasyExcelBuilder buildHeadColor(IndexedColors headColor) {
        this.headColor = headColor;
        return this;
    }

    public EasyExcelBuilder buildContentColor(IndexedColors contentColor) {
        this.contentColor = contentColor;
        return this;
    }
    public EasyExcelBuilder build() {
        return this;
    }

    public String getFileName() {
        return fileName;
    }

    public List<String> getSheetNames() {
        return sheetNames;
    }

    public List<List<String>> getDataModels() {
        return dataModels;
    }

    public List<Set<String>> getIncludeColumnFiledNames() {
        return includeColumnFiledNames;
    }

    public List<Set<String>> getExcludeColumnFiledNames() {
        return excludeColumnFiledNames;
    }

    public String getOssPath() {
        return ossPath;
    }

    public String getFilePath() {
        return filePath;
    }

    public IndexedColors getHeadColor() {
        return headColor;
    }

    public IndexedColors getContentColor() {
        return contentColor;
    }
}
