package com.toolsder.epoi.annotation.field;

import org.apache.poi.ss.usermodel.Font;

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
public @interface ExcelCellFont {

    /**
     * 字体
     *
     * @return
     */
    String name() default "宋体";

    /**
     * 字体大小
     *
     * @return
     */
    short heightInPoints() default 11;


    /**
     * 是否为粗体
     *
     * @return
     */
    boolean bold() default false;

    /**
     * 是否为斜体
     *
     * @return
     */
    boolean italic() default false;

    /**
     * 下划线
     *
     * @return
     */
    byte underline() default Font.U_NONE;
}
