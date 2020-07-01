package com.toolsder.epoi;

import com.toolsder.epoi.annotation.field.*;
import com.toolsder.epoi.annotation.type.ExcelSheet;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;

/**
 * created by 猴子请来的逗比 On 2020/7/1
 * <p>
 * 混合的sheet中的list,也可以通过@ExcelSheet注解实现清单的功能
 *
 * @author by 猴子请来的逗比
 */
@Getter
@Setter
@ExcelSheet(name = "清单sheet",inventorySheet = true)
public class MixingSheetList {

    @ExcelListColumn(
            rowHeight = 14.0f,
            nullValue = "/",
            excelCellStyle = @ExcelCellStyle(
                    borderTop = BorderStyle.THIN,
                    borderBottom = BorderStyle.THIN,
                    borderLeft = BorderStyle.THIN,
                    borderRight = BorderStyle.THIN,
                    wrapText = true,
                    excelCellFont = @ExcelCellFont(heightInPoints = 11)
            ),
            excelListTitle = @ExcelListTitle(
                    excelCellStyle = @ExcelCellStyle(borderTop = BorderStyle.THIN,
                            borderBottom = BorderStyle.THIN,
                            borderLeft = BorderStyle.THIN,
                            borderRight = BorderStyle.THIN,
                            wrapText = true,
                            excelCellFont = @ExcelCellFont(bold = true, heightInPoints = 14)),
                    name = "ID号",
                    rowHeight = 57,
                    columnWidth = 8.38f)
    )
    private String id;

    @ExcelListColumn(
            columnOffset = 1,
            rowHeight = 14.0f,
            nullValue = "/",
            excelCellStyle = @ExcelCellStyle(
                    borderTop = BorderStyle.THIN,
                    borderBottom = BorderStyle.THIN,
                    borderLeft = BorderStyle.THIN,
                    borderRight = BorderStyle.THIN,
                    wrapText = true,
                    excelCellFont = @ExcelCellFont(heightInPoints = 11)
            ),
            excelListTitle = @ExcelListTitle(
                    columnOffset = 1,
                    excelCellStyle = @ExcelCellStyle(borderTop = BorderStyle.THIN,
                            borderBottom = BorderStyle.THIN,
                            borderLeft = BorderStyle.THIN,
                            borderRight = BorderStyle.THIN,
                            wrapText = true,
                            excelCellFont = @ExcelCellFont(bold = true, heightInPoints = 14)),
                    name = "内容",
                    rowHeight = 57,
                    columnWidth = 8.38f)
    )
    private String content;

}
