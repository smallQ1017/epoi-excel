package com.toolsder.epoi.annotation.type;

import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;

import java.lang.annotation.*;

/**
 * created by Qk on 2020/3/25
 * <p>
 * 工作表打印注解
 *
 * @author by 猴子请来的逗比
 */
@Documented
@Target(ElementType.TYPE)
@Inherited
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelPrint {



    /**
     * 打印方向
     * <p>
     * false(默认):纵向，true：横向
     *
     * @return
     */
    boolean landscape() default false;

    /**
     * 纸张格式
     *
     * @return
     */
    short paperSize() default PrintSetup.A4_EXTRA_PAPERSIZE;

    /**
     * 缩放
     *
     * @return
     */
    short scale() default 100;

    /**
     * 重置分辨率
     *
     * @return
     */
    short vResolution() default 300;
}
