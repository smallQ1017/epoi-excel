package com.toolsder.epoi;


import com.toolsder.epoi.annotation.field.*;
import com.toolsder.epoi.annotation.type.ExcelPrint;
import com.toolsder.epoi.annotation.type.ExcelSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by QK on 2020/5/18
 * <p>
 * Excel工具类
 *
 * @author 猴子请来的逗逼
 */
public class ExcelUtils {


    /**
     * 设置工作表参数
     *
     * @param XSSFSheet 工作表
     */
    public static void setSheetParam(XSSFSheet XSSFSheet, ExcelSheet excelSheetAnnotation) {
        XSSFSheet.setMargin(Sheet.TopMargin, excelSheetAnnotation.topMargin());
        XSSFSheet.setMargin(Sheet.BottomMargin, excelSheetAnnotation.bottomMargin());
        XSSFSheet.setMargin(Sheet.LeftMargin, excelSheetAnnotation.leftMargin());
        XSSFSheet.setMargin(Sheet.RightMargin, excelSheetAnnotation.rightMargin());
        XSSFSheet.setHorizontallyCenter(excelSheetAnnotation.horizontallyCenter());
        XSSFSheet.setVerticallyCenter(excelSheetAnnotation.verticallyCenter());
    }

    /**
     * 设置工作表打印参数
     *
     * @param XSSFSheet 工作表
     */
    public static void setPrintParam(XSSFSheet XSSFSheet, ExcelPrint excelPrintAnnotation) {
        XSSFPrintSetup printSetup = XSSFSheet.getPrintSetup();
        printSetup.setPaperSize(excelPrintAnnotation.paperSize());
        printSetup.setLandscape(excelPrintAnnotation.landscape());
        printSetup.setScale(excelPrintAnnotation.scale());
        printSetup.setVResolution(excelPrintAnnotation.vResolution());
    }


    /**
     * 设置字段注解为ExcelCell的单元格内容
     *
     * @param sheet
     * @param value
     * @param excelCellAnnotation
     */
    public static void setExcelCell(XSSFWorkbook workbook, XSSFSheet sheet, Object value, ExcelCell excelCellAnnotation) {
        //行高
        float rowHeight = excelCellAnnotation.rowHeight();
        //列宽
        float columnWidth = excelCellAnnotation.columnWidth();
        //行index
        int rowIndex = excelCellAnnotation.rowIndex();
        //列index
        int columnIndex = excelCellAnnotation.columnIndex();
        //创建单元格
        XSSFCell cell = ExcelUtils.createCell(workbook, sheet, rowIndex, columnIndex, rowHeight, columnWidth, excelCellAnnotation.excelCellMerged(), excelCellAnnotation.excelCellStyle());
        //设置单元格值
        if (null == value) {
            cell.setCellValue(excelCellAnnotation.nullValue());
        } else {
            if (value instanceof Boolean) {
                Boolean valueTemp = (Boolean) value;
                if (valueTemp) {
                    cell.setCellValue(excelCellAnnotation.booleanTrueValue());
                } else {
                    cell.setCellValue(excelCellAnnotation.booleanFalseValue());
                }
            } else {
                cell.setCellValue(String.valueOf(value));
            }
        }
    }

    /**
     * 设置表格图片
     *
     * @param workbook
     * @param sheet
     * @param bytes
     * @param excelPictureAnnotation
     */
    public static void setExcelPicture(XSSFWorkbook workbook, XSSFSheet sheet, byte[] bytes, ExcelPicture excelPictureAnnotation) {
//        CreationHelper helper = workbook.getCreationHelper();
//
//        int pictureIdx = workbook.addPicture(bytes, excelPictureAnnotation.pictureType());
//        //create drawing
//        Drawing<?> drawing = sheet.createDrawingPatriarch();
//        ClientAnchor anchor = helper.createClientAnchor();
//        anchor.setCol1(excelPictureAnnotation.colIndex1());
//        anchor.setRow1(excelPictureAnnotation.rowIndex1());
//        anchor.setCol2(excelPictureAnnotation.colIndex2());
//        anchor.setRow2(excelPictureAnnotation.rowIndex2());
//        anchor.setDx1(excelPictureAnnotation.dx1()* Units.EMU_PER_PIXEL);
//        anchor.setDy1(excelPictureAnnotation.dy1()* Units.EMU_PER_PIXEL);
//        anchor.setDx2(excelPictureAnnotation.dx2()* Units.EMU_PER_PIXEL);
//        anchor.setDy2(excelPictureAnnotation.dy2()* Units.EMU_PER_PIXEL);
//        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
//
//        Picture pict = drawing.createPicture(anchor, pictureIdx);
//
//        //auto-size picture
//        pict.resize();
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor clientAnchor = drawing.createAnchor(
                excelPictureAnnotation.dx1()* Units.EMU_PER_PIXEL,
                excelPictureAnnotation.dy1()* Units.EMU_PER_PIXEL,
                excelPictureAnnotation.dx2()* Units.EMU_PER_PIXEL,
                excelPictureAnnotation.dy2()* Units.EMU_PER_PIXEL,
                excelPictureAnnotation.colIndex1(),
                excelPictureAnnotation.rowIndex1(),
                excelPictureAnnotation.colIndex2(),
                excelPictureAnnotation.rowIndex2()
        );
        drawing.createPicture(clientAnchor, workbook.addPicture(bytes, excelPictureAnnotation.pictureType()));
    }

    /**
     * 设置字段注解为ExcelCell的单元格内容
     *
     * @param sheet
     * @param field
     * @param excelListAnnotation
     */
    public static void setExcelList(Object object, XSSFWorkbook workbook, XSSFSheet sheet, Field field, ExcelList excelListAnnotation) {
        //获取起始列与起始行
        int startRowIndex = excelListAnnotation.startRowIndex();
        int startColumnIndex = excelListAnnotation.startColumnIndex();
        //开始行
        int rowIndex = startRowIndex + 1;
        try {
            field.setAccessible(true);
            if (!(field.get(object) instanceof List)) throw new RuntimeException("字段注解为@ExcelList，但该字段不是List类型。");
            List<?> list = (List<?>) field.get(object);
            for (int i = 0, l = list.size(); i < l; i++) {
                Class<?> tempClz = list.get(i).getClass();
                List<Field> listObjFields = new ArrayList<>();
                while (null != tempClz) {
                    listObjFields.addAll(Arrays.asList(tempClz.getDeclaredFields()));
                    tempClz = tempClz.getSuperclass(); //得到父类,然后赋给自己
                }
                for (Field listObjField : listObjFields) {
                    //获取属性字段的注解
                    Annotation[] fieldAnnotations = listObjField.getAnnotations();
                    for (Annotation fieldAnnotation : fieldAnnotations) {
                        if (fieldAnnotation instanceof ExcelListColumn) {
                            if (i == 0) {
                                setExcelListColumnTitle(workbook, sheet, startRowIndex, startColumnIndex, ((ExcelListColumn) fieldAnnotation).excelListTitle());
                            }
                            //获取字段内容
                            Object value = null;
                            try {
                                listObjField.setAccessible(true);
                                value = listObjField.get(list.get(i));
                            } catch (IllegalAccessException e) {
                                e.printStackTrace();
                            }
                            setExcelListColumn(workbook, sheet, rowIndex, startColumnIndex, value, (ExcelListColumn) fieldAnnotation);
                        }
                    }
                }
                rowIndex++;
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }


    /**
     * 设置List字段的标题
     *
     * @param sheet
     * @param rowIndex
     * @param columnIndex
     * @param listTitleAnnotation
     */
    public static void setExcelListColumnTitle(XSSFWorkbook workbook, XSSFSheet sheet, int rowIndex, int columnIndex, ExcelListTitle listTitleAnnotation) {
        String name = listTitleAnnotation.name();
        int rowOffset = listTitleAnnotation.rowOffset();
        int columnOffset = listTitleAnnotation.columnOffset();
        float rowHeight = listTitleAnnotation.rowHeight();
        float columnWidth = listTitleAnnotation.columnWidth();
        //创建单元格
        XSSFCell cell = ExcelUtils.createCell(workbook, sheet, rowIndex + rowOffset, columnIndex + columnOffset, rowHeight, columnWidth, listTitleAnnotation.excelCellMerged(), listTitleAnnotation.excelCellStyle());
        cell.setCellValue(name);
    }

    /**
     * 设置List字段的列内容
     *
     * @param sheet
     * @param rowIndex
     * @param columnIndex
     * @param value
     * @param excelListColumn
     */
    public static void setExcelListColumn(XSSFWorkbook workbook, XSSFSheet sheet, int rowIndex, int columnIndex, Object value, ExcelListColumn excelListColumn) {
        int rowOffset = excelListColumn.rowOffset();
        int columnOffset = excelListColumn.columnOffset();
        float rowHeight = excelListColumn.rowHeight();
        //创建单元格
        XSSFCell cell = ExcelUtils.createCell(workbook, sheet, rowIndex + rowOffset, columnIndex + columnOffset, rowHeight, excelListColumn.excelCellMerged(), excelListColumn.excelCellStyle());
        //设置单元格值
        if (null == value) {
            cell.setCellValue(excelListColumn.nullValue());
        } else {
            if (value instanceof Boolean) {
                Boolean valueTemp = (Boolean) value;
                if (valueTemp) {
                    cell.setCellValue(excelListColumn.booleanTrueValue());
                } else {
                    cell.setCellValue(excelListColumn.booleanFalseValue());
                }
            } else {
                cell.setCellValue(String.valueOf(value));
            }
        }
    }


    /**
     * 创建单元格并设置单元格所在列的列宽
     *
     * @param workbook             工作薄
     * @param sheet                工作表
     * @param rowIndex             行index
     * @param columnIndex          列index
     * @param rowHeight            行高
     * @param columnWidth          列宽
     * @param cellMergedAnnotation 单元格合并注解
     * @param cellStyleAnnotation  单元格样式注解
     * @return 创建好的单元格
     */
    public static XSSFCell createCell(XSSFWorkbook workbook, XSSFSheet sheet, int rowIndex, int columnIndex, float rowHeight, float columnWidth, ExcelCellMerged cellMergedAnnotation, ExcelCellStyle cellStyleAnnotation) {
        XSSFCell cell = ExcelUtils.createCell(workbook, sheet, rowIndex, columnIndex, rowHeight, cellMergedAnnotation, cellStyleAnnotation);
        sheet.setColumnWidth(columnIndex, (int) ((columnWidth * 8 + 5) / 8 * 265));
        return cell;
    }

    /**
     * 创建单元格，并设置单元格样式
     *
     * @param workbook             工作薄
     * @param sheet                工作表
     * @param rowIndex             行index
     * @param columnIndex          列index
     * @param rowHeight            行高
     * @param cellMergedAnnotation 单元格合并注解
     * @param cellStyleAnnotation  单元格样式注解
     * @return 创建好的单元格
     */
    public static XSSFCell createCell(XSSFWorkbook workbook, XSSFSheet sheet, int rowIndex, int columnIndex, float rowHeight, ExcelCellMerged cellMergedAnnotation, ExcelCellStyle cellStyleAnnotation) {
        //进行单元格合并
        ExcelUtils.mergedCell(sheet, rowIndex, columnIndex, cellMergedAnnotation);
        //创建单元格
        XSSFRow row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
            row.setHeightInPoints(rowHeight);
        }
        XSSFCell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        //设置样式，对合并的所有单元格进行样式设计
        ExcelUtils.setCellStyle(workbook, cell, cellStyleAnnotation);
        //获取合并信息
        int rowCount = cellMergedAnnotation.rowCount();
        int columnCount = cellMergedAnnotation.columnCount();
        if (rowCount > 1) {
            for (int i = rowIndex + 1; i < rowCount + rowIndex; i++) {
                XSSFRow rowTemp = sheet.getRow(i);
                if (rowTemp == null) {
                    rowTemp = sheet.createRow(i);
                }
                for (int a = columnIndex + 1; a < columnCount + columnIndex; a++) {
                    XSSFCell cellTemp = rowTemp.getCell(a);
                    if (cellTemp == null) {
                        cellTemp = row.createCell(a);
                    }
                    //设置样式，对合并的所有单元格进行样式设计
                    ExcelUtils.setCellStyle(workbook, cellTemp, cellStyleAnnotation);
                }
            }
        } else {
            for (int a = columnIndex + 1; a < columnCount + columnIndex; a++) {
                XSSFCell cellTemp = row.getCell(a);
                if (cellTemp == null) {
                    cellTemp = row.createCell(a);
                }
                //设置样式，对合并的所有单元格进行样式设计
                ExcelUtils.setCellStyle(workbook, cellTemp, cellStyleAnnotation);
            }
        }
        return cell;
    }

    /**
     * 单元格样式设置
     *
     * @param workbook        工作薄
     * @param cell            单元格
     * @param styleAnnotation 样式注解
     */
    public static void setCellStyle(XSSFWorkbook workbook, XSSFCell cell, ExcelCellStyle styleAnnotation) {
        CellStyle cellStyle = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        //设置字体
        ExcelCellFont fontAnnotation = styleAnnotation.excelCellFont();
        font.setBold(fontAnnotation.bold());
        font.setItalic(fontAnnotation.italic());
        font.setFontName(fontAnnotation.name());
        font.setFontHeightInPoints(fontAnnotation.heightInPoints());
        font.setUnderline(fontAnnotation.underline());
        //设置样式

        cellStyle.setFont(font);
        cellStyle.setWrapText(styleAnnotation.wrapText());
        cellStyle.setAlignment(styleAnnotation.horizontalAlignment());
        cellStyle.setVerticalAlignment(styleAnnotation.verticalAlignment());
        cellStyle.setBorderLeft(styleAnnotation.borderLeft());
        cellStyle.setBorderRight(styleAnnotation.borderRight());
        cellStyle.setBorderTop(styleAnnotation.borderTop());
        cellStyle.setBorderBottom(styleAnnotation.borderBottom());
        cell.setCellStyle(cellStyle);

    }

    /**
     * 单元格合并，如果注解描述所需合并的行数为1且列数为1则不进行合并
     *
     * @param sheet            工作表
     * @param rowIndex         行index
     * @param columnIndex      列index
     * @param mergedAnnotation 合并注解
     */
    public static void mergedCell(XSSFSheet sheet, int rowIndex, int columnIndex, ExcelCellMerged mergedAnnotation) {
        //获取合并信息
        int rowCount = mergedAnnotation.rowCount();
        int columnCount = mergedAnnotation.columnCount();
        //如果需要合并则开始合并
        if (!(rowCount == 1 && columnCount == 1))
            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + rowCount - 1, columnIndex, columnIndex + columnCount - 1));
    }
}
