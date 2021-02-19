package com.toolsder.epoi;

import com.google.common.base.Strings;
import com.google.common.collect.ObjectArrays;
import com.toolsder.epoi.annotation.field.ExcelCell;
import com.toolsder.epoi.annotation.field.ExcelList;
import com.toolsder.epoi.annotation.field.ExcelListColumn;
import com.toolsder.epoi.annotation.field.ExcelPicture;
import com.toolsder.epoi.annotation.type.ExcelSheet;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;

/**
 * created by 猴子请来的逗比 On 2020/12/14
 *
 * @author by 猴子请来的逗比
 */
public class ExcelFileToObject {

    /**
     * 定义工作表的注解
     */
    private ExcelSheet excelSheetAnnotation;

    /**
     * 实体对象的class
     */
    private Class<?> clz;


    private XSSFSheet xssfSheet;


    public ExcelFileToObject(Class<?> clz, XSSFWorkbook workbook) {
        this.clz = clz;
        this.xssfSheet = workbook.getSheetAt(0);
        this.extractClass();
    }

    /**
     * 获取清单工作表
     *
     * @return 列表
     * @throws InstantiationException
     * @throws IllegalAccessException
     */
    public List generateList() throws InstantiationException, IllegalAccessException {
        List temp = new ArrayList();
        if (!excelSheetAnnotation.inventorySheet())
            throw new RuntimeException("类的@ExcelSheet注解中的inventorySheet属性为false,请设置为true");
        //开始设置数据
        this.disposeInventoryField(1, 0, 0, temp, this.clz);
        return temp;
    }

    /**
     * 获取对象工作表
     *
     * @return 对象
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    public Object generateObject() throws IllegalAccessException, InstantiationException, ClassNotFoundException {
        Object obj = clz.newInstance();
        if (excelSheetAnnotation.inventorySheet())
            throw new RuntimeException("类的@ExcelSheet注解中的inventorySheet属性为true,请设置为false");
        Field[] fields = clz.getDeclaredFields();
        for (Field field : fields) {
            //设置属性可访问
            field.setAccessible(true);
            //获取属性字段的注解
            Annotation[] fieldAnnotations = field.getAnnotations();
            //处理字段属性注解
            for (Annotation annotation : fieldAnnotations) {
                if (annotation instanceof ExcelCell) {
                    ExcelCell temp = (ExcelCell) annotation;
                    //开始设置数据
                    XSSFCell xssfCell = xssfSheet.getRow(temp.rowIndex()).getCell(temp.columnIndex());
                    xssfCell.setCellType(CellType.STRING);
                    field.set(obj, xssfCell.getStringCellValue());
                } else if (annotation instanceof ExcelList) {
                    ExcelList annotationTemp = (ExcelList) annotation;
                    //获取list中的class
                    ParameterizedType t = (ParameterizedType) field.getGenericType();
                    Type[] actualTypeArguments = t.getActualTypeArguments();
                    if (actualTypeArguments.length == 1) {
                        Type actualTypeArgument = actualTypeArguments[0];
                        List temp = new ArrayList();
                        this.disposeInventoryField(annotationTemp.startRowIndex() + 1, annotationTemp.rowCount(), annotationTemp.startColumnIndex(), temp, Class.forName(actualTypeArgument.getTypeName()));
                        field.set(obj, temp);
                    }
                } else if (annotation instanceof ExcelPicture) {

                }
            }
        }
        return obj;
    }

//
//    /**
//     * 生成工作表
//     */
//    public Object generateObject() throws InstantiationException, IllegalAccessException {
//        //判断所要创建的工作表是否为清单工作表
//        if (!excelSheetAnnotation.inventorySheet()) {
//            //开始处理字段
//            Field[] fields = clz.getDeclaredFields();
//            for (Field field : fields) {
//                this.disposeField(field);
//            }
//        } else {
//
//        }
//        return obj;
//    }

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
        if (null == excelSheetAnnotation) throw new RuntimeException("请定义ExcelSheet注解后在使用");
    }


    /**
     * 开始处理字段
     *
     * @param field 对象类字段
     */
    private void disposeField(Field field) {
        //获取属性字段的注解
        Annotation[] fieldAnnotations = field.getAnnotations();
        //处理字段属性注解
        for (Annotation annotation : fieldAnnotations) {
            if (annotation instanceof ExcelCell) {

            } else if (annotation instanceof ExcelList) {

            } else if (annotation instanceof ExcelPicture) {

            }
        }
    }

    /**
     * 处理清单式表格字段
     *
     * @param
     */
    private void disposeInventoryField(int startRowIndex, int rowCount, int startColumnIndex, List temp, Class clz) throws IllegalAccessException, InstantiationException {
        int endRowIndex = rowCount == 0 ? this.xssfSheet.getLastRowNum() + 1 : rowCount + startRowIndex;
        //获取字段
        //获取对象的属性字段，包括父类
        Class<?> supperClz = clz.getSuperclass();
        Field[] fields = clz.getDeclaredFields();
        while (null != supperClz) {
            fields = ObjectArrays.concat(fields, supperClz.getDeclaredFields(), Field.class);
            supperClz = supperClz.getSuperclass(); //得到父类,然后赋给自己
        }
        for (int i = startRowIndex; i < endRowIndex; i++) {
            //获取行内容
            XSSFRow row = this.xssfSheet.getRow(i);
            if (null != row) {
                //创建对象
                Object object = clz.newInstance();
                //循环获取列内容
                for (Field listObjField : fields) {
                    //设置属性可访问
                    listObjField.setAccessible(true);
                    //获取属性字段的注解
                    Annotation[] fieldAnnotations = listObjField.getAnnotations();
                    for (Annotation fieldAnnotation : fieldAnnotations) {
                        if (fieldAnnotation instanceof ExcelListColumn) {

                            //开始设置数据
                            XSSFCell xssfCell = row.getCell(startColumnIndex + ((ExcelListColumn) fieldAnnotation).columnOffset());
                            if (null != xssfCell) {
                                xssfCell.setCellType(CellType.STRING);
                                if (Strings.isNullOrEmpty(xssfCell.getStringCellValue())) {
                                    listObjField.set(object, null);
                                } else {
                                    listObjField.set(object, xssfCell.getStringCellValue());
                                }
                            }

                        }
                    }
                }
                temp.add(object);
            }
        }
    }
}
