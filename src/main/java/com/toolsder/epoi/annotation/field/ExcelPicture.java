package com.toolsder.epoi.annotation.field;

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
    int col1() default 0;

    /**
     * 开始单元格行标
     */
    int row1() default 0;

    /**
     * 结束单元格列标
     */
    int col2() default 0;

    /**
     * 结束单元格行标
     */
    int row2() default 0;

    /**
     * 开始单元格x坐标
     */
    int dx1() default 0;

    /**
     * 开始单元格y坐标
     */
    int dy1() default 0;

    /**
     * 结束单元格x坐标
     */
    int dx2() default 0;

    /**
     * 结束单元格y坐标
     */
    int dy2() default 0;


}
