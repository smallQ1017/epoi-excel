package com.toolsder.epoi.annotation.field;

import java.lang.annotation.*;

/**
 * created by Qk on 2020/3/22
 *
 * @author by 猴子请来的逗比
 */
@Documented
@Target(ElementType.FIELD)
@Inherited
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelListColumn {
    /**
     * 起始行偏移
     *
     * @return
     */
    int rowOffset() default 0;

    /**
     * 起始列偏移
     *
     * @return
     */
    int columnOffset() default 0;

    /**
     * 行高 以像素为单位
     *
     * @return
     */
    float rowHeight() default 13.5f;

    /**
     * null值
     *
     * @return
     */
    String nullValue() default "";

    /**
     * true值
     *
     * @return
     */
    String booleanTrueValue() default "";

    /**
     * false值
     *
     * @return
     */
    String booleanFalseValue() default "";

    /**
     * 列标题
     *
     * @return
     */
    ExcelListTitle excelListTitle() default @ExcelListTitle;


    /**
     * 单元格样式
     *
     * @return
     */
    ExcelCellStyle excelCellStyle() default @ExcelCellStyle;

    /**
     * 单元格合并
     *
     * @return
     */
    ExcelCellMerged excelCellMerged() default @ExcelCellMerged;
}
