package com.toolsder.epoi.annotation.field;

import org.apache.poi.ss.usermodel.Workbook;

import java.lang.annotation.*;

/**
 * created by Qk on 2020/3/22
 * excel表格的图片
 *
 * @author by 猴子请来的逗比
 */
@Documented
@Target(ElementType.FIELD)
@Inherited
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelPicture {
    /**
     * 开始单元格列标
     */
    int colIndex1() default 0;

    /**
     * 开始单元格行标
     */
    int rowIndex1() default 0;

    /**
     * 结束单元格列标
     */
    int colIndex2() default 0;

    /**
     * 结束单元格行标
     */
    int rowIndex2() default 0;

    /**
     * 开始单元格x坐标(单位为单元格像素值)
     */
    int dx1() default 0;

    /**
     * 开始单元格y坐标(单位为单元格像素值)
     */
    int dy1() default 0;

    /**
     * 结束单元格x坐标(单位为单元格像素值)
     * 图片右下角在第二个单元格里面的x位置
     */
    int dx2() default 0;

    /**
     * 结束单元格y坐标(单位为单元格像素值)
     * 图片右下角在第二个单元格里面的y位置
     */
    int dy2() default 0;


    /**
     * 图片类型
     */
    int pictureType() default Workbook.PICTURE_TYPE_PNG;

}