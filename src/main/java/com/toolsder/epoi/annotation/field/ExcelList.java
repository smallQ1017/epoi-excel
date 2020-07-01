package com.toolsder.epoi.annotation.field;

import java.lang.annotation.*;

/**
 * created by Qk on 2020/3/25
 * <p>
 * 标记表格对象中的list字段，用于
 *
 * @author by 猴子请来的逗比
 */
@Documented
@Target(ElementType.FIELD)
@Inherited
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelList {

    /**
     * list数据起始row的index
     *
     * @return
     */
    int startRowIndex() default 0;


    /**
     * list数据起始column的index
     *
     * @return
     */
    int startColumnIndex() default 0;

}
