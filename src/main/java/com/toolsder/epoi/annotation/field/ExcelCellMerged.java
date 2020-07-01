package com.toolsder.epoi.annotation.field;

import java.lang.annotation.*;

/**
 * created by Qk on 2020/3/22
 * excel 单元格合并注解
 *
 * @author by 猴子请来的逗比
 */
@Documented
@Target(ElementType.ANNOTATION_TYPE)
@Inherited
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCellMerged {
    /**
     * 合并行数
     *
     * @return
     */
    int rowCount() default 1;

    /**
     * 合并列数
     *
     * @return
     */
    int columnCount() default 1;
}
