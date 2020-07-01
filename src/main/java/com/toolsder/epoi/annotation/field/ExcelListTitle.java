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
public @interface ExcelListTitle {

    /**
     * 标题名称
     *
     * @return
     */
    String name() default "标题";

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
     * 行高
     *
     * @return
     */
    float rowHeight() default 13.5f;

    /**
     * 列宽
     *
     * @return
     */
    float columnWidth() default 8.38f;

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
