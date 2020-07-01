package com.toolsder.epoi.annotation.field;

import java.lang.annotation.*;

/**
 * created by Qk on 2020/3/22
 * excel表格的列注解
 *
 * @author by 猴子请来的逗比
 */
@Documented
@Target(ElementType.FIELD)
@Inherited
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCell {
    /**
     * 行高
     */
    float rowHeight() default 13.5f;

    /**
     * 列宽
     */
    float columnWidth() default 8.38f;

    /**
     * 行index
     */
    int rowIndex() default 0;

    /**
     * 列index
     */
    int columnIndex() default 0;

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
