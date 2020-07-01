package com.toolsder.epoi.annotation.type;

import java.lang.annotation.*;

/**
 * created by Qk on 2020/3/25
 * <p>
 * 工作表注解
 *
 * @author by 猴子请来的逗比
 */
@Documented
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSheet {

    /**
     * 是否清单工作表
     */
    boolean inventorySheet() default false;
    /**
     * 工作表名称
     */
    String name() default "sheet";


    /**
     * 水平居中
     */
    boolean horizontallyCenter() default true;

    /**
     * 重置居中
     */
    boolean verticallyCenter() default true;

    /**
     * 页面上边距 以英寸为单位
     */
    double topMargin() default 0.984252;

    /**
     * 页面下边距 以英寸为单位
     */
    double bottomMargin() default 0.984252;

    /**
     * 页面左边距 以英寸为单位
     */
    double leftMargin() default 0.7480315;

    /**
     * 页面右边距 以英寸为单位
     */
    double rightMargin() default 0.7480315;
}
