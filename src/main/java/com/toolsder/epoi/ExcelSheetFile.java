package com.toolsder.epoi;

import com.google.common.collect.ObjectArrays;

import com.toolsder.epoi.annotation.field.ExcelCell;
import com.toolsder.epoi.annotation.field.ExcelList;
import com.toolsder.epoi.annotation.field.ExcelListColumn;
import com.toolsder.epoi.annotation.field.ExcelPicture;
import com.toolsder.epoi.annotation.type.ExcelPrint;
import com.toolsder.epoi.annotation.type.ExcelSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.List;

/**
 * Created by QK on 2020/5/14
 * 将对象制作成Excel表格
 *
 * @author 猴子请来的逗逼
 */
public class ExcelSheetFile {

    /**
     * 定义工作表的注解
     */
    private ExcelSheet excelSheetAnnotation;

    /**
     * 定义工作表打印配置的注解
     */
    private ExcelPrint excelPrintAnnotation;

    /**
     * 实体对象
     */
    private Object object;

    /**
     * 实体对象的class
     */
    private Class<?> clz;

    /**
     * 保存实体对象数据的工作薄
     */
    private XSSFWorkbook workbook;

    public ExcelSheetFile(Object object, XSSFWorkbook workbook) {
        //设置对象
        this.setObject(object, workbook);
    }

    /**
     * 生成工作表
     */
    public void generateSheet() {
        XSSFSheet sheet = this.createSheet();
        //判断所要创建的工作表是否为清单工作表
        if (!excelSheetAnnotation.inventorySheet()) {
            //开始处理字段
            Field[] fields = clz.getDeclaredFields();
            for (Field field : fields) {
                this.disposeField(sheet, field);
            }
        } else {
            if (!(this.object instanceof List))
                throw new RuntimeException("当注解@ExcelSheet中的inventorySheet属性为true时，请确保传入的对象为List");
            //判断对象是否为collection
            List<?> objects = (List<?>) object;
            //获取对象的属性字段，包括父类
            Class<?> supperClz = clz.getSuperclass();
            Field[] fields = clz.getDeclaredFields();
            while (null != supperClz) {
                fields = ObjectArrays.concat(fields, supperClz.getDeclaredFields(), Field.class);
                supperClz = supperClz.getSuperclass(); //得到父类,然后赋给自己
            }
            this.disposeInventoryField(sheet, objects, fields);
        }
    }

    /**
     * 设置需要转成工作表的对象
     *
     * @param object 要转成工作表的对象
     */
    private void setObject(Object object, XSSFWorkbook workbook) {
        if (object instanceof List) {
            List<Object> objects = (List<Object>) object;
            if (objects.size() <= 0) throw new RuntimeException("请勿传入空list");
            //获取对象的Class
            this.clz = objects.get(0).getClass();
        } else {
            //获取对象的Class
            this.clz = object.getClass();
        }
        this.object = object;
        this.extractClass();
        this.workbook = workbook;
    }

    /**
     * 提取类的注解
     * <code>@ExcelSheet</code>
     * <code>@ExcelPrint</code>
     *
     * @throws RuntimeException 如果没有进行@ExcelSheet注解则抛出异常
     */
    private void extractClass() {

        //获取对象的ExcelSheet注解
        excelSheetAnnotation = clz.getAnnotation(ExcelSheet.class);
        excelPrintAnnotation = clz.getAnnotation(ExcelPrint.class);
        if (null == excelSheetAnnotation) throw new RuntimeException("请定义ExcelSheet注解后在使用");
    }

    /**
     * 创建一个工作表
     *
     * @return 一个工作表
     */
    private XSSFSheet createSheet() {
        String name = excelSheetAnnotation.name();
        int count = 2;
        while (this.workbook.getSheetIndex(name) >= 0) {
            name = excelSheetAnnotation.name() + count;
            count++;
        }
        XSSFSheet sheet = this.workbook.createSheet(name);
        ExcelUtils.setSheetParam(sheet, excelSheetAnnotation);
        if (null != excelPrintAnnotation) {
            ExcelUtils.setPrintParam(sheet, excelPrintAnnotation);
        }
        return sheet;
    }

    /**
     * 开始处理字段
     *
     * @param field 对象类字段
     */
    private void disposeField(XSSFSheet heet, Field field) {
        //获取属性字段的注解
        Annotation[] fieldAnnotations = field.getAnnotations();
        //处理字段属性注解
        for (Annotation annotation : fieldAnnotations) {
            if (annotation instanceof ExcelCell) {
                //获取字段内容
                Object value = null;
                try {
                    field.setAccessible(true);
                    value = field.get(object);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
                ExcelUtils.setExcelCell(workbook, heet, value, (ExcelCell) annotation);
            } else if (annotation instanceof ExcelList) {
                ExcelUtils.setExcelList(object, workbook, heet, field, (ExcelList) annotation);
            } else if (annotation instanceof ExcelPicture) {
                //获取字段内容
                Object value = null;
                try {
                    field.setAccessible(true);
                    value = field.get(object);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
                if (value instanceof byte[]) {
                    ExcelUtils.setExcelPicture(workbook, heet, (byte[]) value, (ExcelPicture) annotation);
                } else {
                    throw new RuntimeException("请确认数据为byte[]");
                }
            }
        }
    }

    /**
     * 处理清单式表格字段
     *
     * @param heet
     * @param objects
     * @param fields
     */
    private void disposeInventoryField(XSSFSheet heet, List<?> objects, Field[] fields) {
        //定义内容起始行
        int rowIndex = 1;
        for (int i = 0, l = objects.size(); i < l; i++) {
            for (Field listObjField : fields) {
                //获取属性字段的注解
                Annotation[] fieldAnnotations = listObjField.getAnnotations();
                for (Annotation fieldAnnotation : fieldAnnotations) {
                    if (fieldAnnotation instanceof ExcelListColumn) {
                        if (i == 0) {
                            ExcelUtils.setExcelListColumnTitle(workbook, heet, 0, 0, ((ExcelListColumn) fieldAnnotation).excelListTitle());
                        }
                        //获取字段内容
                        Object value = null;
                        try {
                            listObjField.setAccessible(true);
                            value = listObjField.get(objects.get(i));
                        } catch (IllegalAccessException e) {
                            e.printStackTrace();
                        }
                        ExcelUtils.setExcelListColumn(workbook, heet, rowIndex, 0, value, (ExcelListColumn) fieldAnnotation);
                    }
                }
            }
            rowIndex++;
        }
    }
}
