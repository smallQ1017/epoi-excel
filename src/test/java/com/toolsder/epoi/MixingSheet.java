package com.toolsder.epoi;

import com.toolsder.epoi.annotation.field.*;
import com.toolsder.epoi.annotation.type.ExcelSheet;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * created by 猴子请来的逗比 On 2020/7/1
 * <p>
 * 一个混合的sheet定义类
 *
 * @author by 猴子请来的逗比
 */
@ExcelSheet(name = "混合sheet")
@Getter
@Setter
public class MixingSheet {
    /**
     * 一个长合并标题
     */
//    @ExcelCell(
//            rowIndex = 0,
//            columnIndex = 0,
//            rowHeight = 20.0f,
//            nullValue = "/",
//            excelCellStyle = @ExcelCellStyle(
//                    borderLeft = BorderStyle.THIN,
//                    borderTop = BorderStyle.DASH_DOT_DOT,
//                    borderRight = BorderStyle.HAIR,
//                    borderBottom = BorderStyle.MEDIUM_DASH_DOT,
//                    wrapText = true,
//                    excelCellFont = @ExcelCellFont(
//                            name = "微软雅黑",
//                            heightInPoints = 20,
//                            bold = true,
//                            italic = true,
//                            underline = Font.U_DOUBLE_ACCOUNTING
//                    )
//            ),
//            excelCellMerged = @ExcelCellMerged(
//                    rowCount = 2,
//                    columnCount = 2
//            )
//    )
    private String title;

    /**
     * 一个非合并小标题
     */
//    @ExcelCell(
//            rowIndex = 2,
//            columnIndex = 0,
//            rowHeight = 21.0f,
//            nullValue = "/",
//            excelCellStyle = @ExcelCellStyle(
//                    borderLeft = BorderStyle.THIN,
//                    borderTop = BorderStyle.DASH_DOT_DOT,
//                    borderRight = BorderStyle.HAIR,
//                    borderBottom = BorderStyle.MEDIUM_DASH_DOT,
//                    wrapText = true,
//                    excelCellFont = @ExcelCellFont(
//                            heightInPoints = 16,
//                            bold = true,
//                            italic = true,
//                            underline = Font.U_DOUBLE_ACCOUNTING
//                    )
//            )
//    )
    private String subTitle;

    /**
     * 定义列表内容，它从第三行的第二列开始
     */
//    @ExcelList(
//            startRowIndex = 3,
//            startColumnIndex = 1
//    )
    private List<MixingSheetList> mixingSheetLists;


    /**
     * 定义一个图片
     */
    @ExcelPicture(dx1 = 0,dy1 = 0,dx2=50,dy2 = 12,colIndex1 = 0,rowIndex1 = 0,colIndex2 = 0,rowIndex2 = 0)
    private byte[] img;

}
