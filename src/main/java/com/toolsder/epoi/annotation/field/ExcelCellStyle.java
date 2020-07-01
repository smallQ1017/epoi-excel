package com.toolsder.epoi.annotation.field;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

/**
 * created by Qk on 2020/3/22
 * excel表格的列注解
 *
 * @author by 猴子请来的逗比
 */
@Documented
@Target(ElementType.ANNOTATION_TYPE)
@Inherited
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCellStyle {
    /**
     * 水平对齐
     */
    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.CENTER;

    /**
     * 垂直对齐
     */
    VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;


    /**
     * 左边框样式
     */
    BorderStyle borderLeft() default BorderStyle.NONE;

    /**
     * 右边框样式
     */
    BorderStyle borderRight() default BorderStyle.NONE;

    /**
     * 上边框样式
     */
    BorderStyle borderTop() default BorderStyle.NONE;

    /**
     * 下边框样式
     */
    BorderStyle borderBottom() default BorderStyle.NONE;

    /**
     * 文本换行
     */
    boolean wrapText() default false;

    /**
     * 字体
     *
     * @return
     */
    ExcelCellFont excelCellFont() default @ExcelCellFont;
}
