package com.xuxiong.easyexceldemo.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.util.StringUtils;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.List;
import java.util.Set;

/**
 * easyExcel 操作excel 工具类
 * @author 许兄
 * Created on 2019/12/11
 */
public class EasyExcelUtil {

    private static final Logger logger = LoggerFactory.getLogger(EasyExcelUtil.class);
    /**
     * 导出excel简单通用的方法
     * @param easyExcelBuilder  基础参数
     * @param obj  对应的实体类
     * @param data  实体对应的数据集
     * @return
     */
    public static boolean writeExcel(EasyExcelBuilder easyExcelBuilder, Class<?> obj, List<List<?>> data){

        try{

            String filePath =  easyExcelBuilder.getFilePath() + easyExcelBuilder.getFileName();
            ExcelWriter excelWriter = EasyExcel.write(filePath, obj).build();
            int flag = 0;
            List<Set<String>> includeColumnFiledNames = easyExcelBuilder.getIncludeColumnFiledNames();
            if(includeColumnFiledNames != null){

            }
            List<Set<String>> excludeColumnFiledNames = easyExcelBuilder.getExcludeColumnFiledNames();
            for (String sheetName : easyExcelBuilder.getSheetNames()) {
                ExcelWriterSheetBuilder excelWriterSheetBuilder = EasyExcel.writerSheet(flag, sheetName).registerWriteHandler(new LongestMatchColumnWidthStyleStrategy());
                List<?> objects = data.get(flag);
                //导入特定的字段
                if(includeColumnFiledNames != null){
                    if(flag < includeColumnFiledNames.size()){
                        Set<String> includeColumnsName = includeColumnFiledNames.get(flag);
                        excelWriterSheetBuilder.includeColumnFiledNames(includeColumnsName);
                    }
                }else if(excludeColumnFiledNames != null){
                    if(flag < excludeColumnFiledNames.size()){
                        Set<String> excludeColumnName = excludeColumnFiledNames.get(flag);
                        excelWriterSheetBuilder.excludeColumnFiledNames(excludeColumnName);
                    }
                }
                // 头的策略
                WriteCellStyle headWriteCellStyle = new WriteCellStyle();
                // 背景设置
                headWriteCellStyle.setFillForegroundColor(easyExcelBuilder.getHeadColor().index);
                // 内容的策略
                WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
                if(easyExcelBuilder.getContentColor() != null){
                    // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.头默认了 FillPatternType所以可以不指定
                    contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
                    contentWriteCellStyle.setFillForegroundColor(easyExcelBuilder.getContentColor().index);
                }
                // 这个策略是 头是头的样式 内容是内容的样式 其他的策略可以自己实现
                HorizontalCellStyleStrategy horizontalCellStyleStrategy = new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
                excelWriterSheetBuilder.registerWriteHandler(horizontalCellStyleStrategy);
                excelWriter.write(objects, excelWriterSheetBuilder.build());
                flag++;
            }
            /// 千万别忘记finish 会帮忙关闭流
            excelWriter.finish();

            //需要同步到oss
            String ossPath = easyExcelBuilder.getOssPath();
            if(StringUtils.isEmpty(ossPath)){
                if (ossPath.startsWith("/")) {
                    ossPath = ossPath.substring(1);
                }
                ///######## 该处为上传文件的逻辑 上传完文件会进行删除原文件
                //删除本地文件
                try{
                    new File(filePath).delete();
                }catch (Exception ex){
                    logger.error("EasyExcelUtil, deleted file error", ex);
                }
            }
            return true;
        }catch (Exception e){
            logger.error("EasyExcelUtil,writeExcel, error : {}", JSONObject.toJSONString(easyExcelBuilder), e);
        }
        return false;
    }
}
