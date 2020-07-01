# E-POI
该项目主要以通过注解的方式将一个实体类变成一个2007版excel的**sheet**。在生成工作表前你的首先创建一个工作薄（XSSFWorkbook）
----------
## 注解说明
### 类注解
#### @ExcelSheet

当类定义了该注解后即可依据该类定义的注解生成Excel的sheet。当然了，这其中还是需要配合字段的注解才能生成一个比较完整的sheet。

| 参数 | 类型 | 默认值 | 说明 |
| :----: | :----: | :----: | ---- |
| inventorySheet | boolean | false | 用来说明被注解类的对象在生成excel文件时是否为清单文件，所谓清单文件即指如同数据库的一张表 |
| name | String | sheet | 工作表的名称，如果工作表名称出现重复则自动以定义的名称+序号创建新的名称 |
| horizontallyCenter | boolean | true | 设置页面在打印时是否水平居中 |
| verticallyCenter | boolean | true | 设置页面在打印时是否垂直居中 | 
| topMargin | double | 0.984252 | 上边距，以英寸为单位 |
| bottomMargin | double | 0.984252 | 下边距，以英寸为单位 |
| leftMargin | double | 0.7480315 | 左边距，以英寸为单位 |
| rightMargin | double | 0.7480315 | 右边距，以英寸为单位 |

使用方法：
```java
@ExcelSheet(name="测试工作表",leftMargin = 0.1968504, rightMargin = 0.1968504, verticallyCenter = false)
public class TestEntity{}
```
#### @ExcelPrint
该注解主要是为了便于定义工作表的打印参数，该注解单独使用无效，必须配合**@ExcelSheet**一起使用

| 参数 | 类型 | 默认值 | 说明 |
| :----: | :----: | :----: | ---- |
| landscape | boolean | false | 打印方向,false：纵向，true：横向 |
| paperSize | short | PrintSetup.A4_EXTRA_PAPERSIZE | 纸张格式，直接使用poi中PrintSetup接口所定义的一些常量 |
| scale | short | 100 | 缩放比例 |
| vResolution | short | 300 | 垂直分辨率 |

使用方法：
```java
@ExcelSheet(name="测试工作表",leftMargin = 0.1968504, rightMargin = 0.1968504, verticallyCenter = false)
@ExcelPrint(landscape = true, scale = 81)
public class TestEntity{}
```
### 字段注解
#### @ExcelCell
如果定义了@ExcelSheet的类中的字段需要在sheet中显示其内容，则需要定义该注解。该注解定义了一个单元格，该单元格可以是合并的，也可以是非合并的，单元格的内容即为字段的内容。**注意：该注解在父类中生效**

| 参数 | 类型 | 默认值 | 说明 |
| :----: | :----: | :----: | ---- |
| rowHeight | float | 13.5f | 行高。一般如果有多个字段在一行内，则只需给其中一个字段定义行高即可。（注意：如果字段的行定义在同一行内且都定义的行高，则一般以最后一个字段定义的行高为最终行高） |
| columnWidth | float | 8.38f | 列宽，实际列宽可能有细微的差异，但差异不大。一般如果有多个字段在一列内，则只需给其中一个字段定义列宽即可。（注意：如果字段的行定义在同一列下且都定义的列宽，则一般以最后一个字段定义的列宽为最终列宽） |
| rowIndex | int | 0 | 行的index |
| columnIndex | int | 0 | 列的index |
| nullValue | String |  | 当字段值为null时单元格的内容 |
| booleanTrueValue | String |  | 当字段值为boolean型且值为true时单元格的内容 |
| booleanFalseValue | String |  | 当字段值为boolean型且值为false时单元格的内容 |
| excelCellStyle | @ExcelCellStyle | @ExcelCellStyle  | 单元格样式，参见@ExcelCellStyle注解 |
| excelCellMerged | @ExcelCellMerged | @ExcelCellMerged  | 单元格合并方式，参见@ExcelCellMerged注解 |

使用方法：
```java
@ExcelSheet(name="测试工作表",leftMargin = 0.1968504, rightMargin = 0.1968504, verticallyCenter = false)
@ExcelPrint(landscape = true, scale = 81)
public class TestEntity{

    @ExcelCell(
            rowHeight=16.58f,
            columnWidth = 10.5f,
            rowIndex=0,
            columnIndex = 1,
            nullValue="/"
    )
    private String testField;

    @ExcelCell(
            rowHeight=16.58f,
            columnWidth = 10.5f,
            rowIndex=0,
            columnIndex = 1,
            nullValue="/",
            booleanTrueValue = "√",
            booleanFalseValue = "×"
    )
    private Boolean testBooleanField;
}
```
#### @ExcelCellStyle
该注解用来定义单元格的样式。

| 参数 | 类型 | 默认值 | 说明 |
| :----: | :----: | :----: | ---- |
| horizontalAlignment | HorizontalAlignment | HorizontalAlignment.CENTER | 水平对齐方式 |
| verticalAlignment | VerticalAlignment | VerticalAlignment.CENTER | 垂直对齐方式 |
| borderLeft | BorderStyle | BorderStyle.NONE | 左边框样式 |
| borderRight | BorderStyle | BorderStyle.NONE | 右边框样式 |
| borderTop | BorderStyle | BorderStyle.NONE | 上边框样式 |
| borderBottom | BorderStyle | BorderStyle.NONE | 下边框样式 |
| wrapText | boolean | false | 单元格文本是否自动换行。false:不换行，true:换行 |
| excelCellFont | @ExcelCellFont | @ExcelCellFont | 单元格字体，参见@ExcelCellFont注解 |

使用方法：
```java
@ExcelSheet(name = "测试工作表", leftMargin = 0.1968504, rightMargin = 0.1968504, verticallyCenter = false)
@ExcelPrint(landscape = true, scale = 81)
public class TestEntity {

    @ExcelCell(
            rowHeight = 16.58f,
            columnWidth = 10.5f,
            rowIndex = 0,
            columnIndex = 1,
            nullValue = "/",
            excelCellStyle = @ExcelCellStyle(
                    horizontalAlignment = HorizontalAlignment.LEFT,
                    verticalAlignment = VerticalAlignment.TOP,
                    borderTop = BorderStyle.THIN,
                    borderBottom = BorderStyle.THIN,
                    borderLeft = BorderStyle.THIN,
                    borderRight = BorderStyle.THIN,
                    wrapText = true
            )
    )
    private String testField;
}
```
#### @ExcelCellFont
该注解用来定义单元格的字体。

| 参数 | 类型 | 默认值 | 说明 |
| :----: | :----: | :----: | ---- |
| name | String | 宋体 | 字体名称 |
| heightInPoints | short | 11 | 字体大小 |
| bold | boolean | false | 是否为粗体。false:不是，true:是 |
| italic | boolean | false | 是否为斜体。false:不是，true:是 |
| underline | byte | Font.U_NONE | 下划线 |

使用方法：
```java
@ExcelSheet(name = "测试工作表", leftMargin = 0.1968504, rightMargin = 0.1968504, verticallyCenter = false)
@ExcelPrint(landscape = true, scale = 81)
public class TestEntity {

    @ExcelCell(
            rowHeight = 16.58f,
            columnWidth = 10.5f,
            rowIndex = 0,
            columnIndex = 1,
            nullValue = "/",
            excelCellStyle = @ExcelCellStyle(
                    horizontalAlignment = HorizontalAlignment.LEFT,
                    verticalAlignment = VerticalAlignment.TOP,
                    borderTop = BorderStyle.THIN,
                    borderBottom = BorderStyle.THIN,
                    borderLeft = BorderStyle.THIN,
                    borderRight = BorderStyle.THIN,
                    wrapText = true,
                    excelCellFont = @ExcelCellFont(
                            heightInPoints = 10,
                            bold=true,
                            underline = Font.U_DOUBLE
                    )
            )
    )
    private String testField;
}
```
#### @ExcelCellMerged
该注解用来操作单元格的合并。

| 参数 | 类型 | 默认值 | 说明 |
| :----: | :----: | :----: | ---- |
| rowCount | int | 1 | 合并行数 |
| columnCount | int | 1 | 合并列数 |


使用方法：
```java
@ExcelSheet(name = "测试工作表", leftMargin = 0.1968504, rightMargin = 0.1968504, verticallyCenter = false)
@ExcelPrint(landscape = true, scale = 81)
public class TestEntity {

    @ExcelCell(
            excelCellMerged = @ExcelCellMerged(
                    columnCount = 20
            )
    )
    private String testField;
}
```
#### @ExcelList
当@ExcelSheet定义的类中有类型为List的字段且需要将该List字段中的内容显示到表格中则需对该字段定义该注解，并且List的T中的字段需要用@ExcelListColumn进行定义。

| 参数 | 类型 | 默认值 | 说明 |
| :----: | :----: | :----: | ---- |
| startRowIndex | int | 0 | list内容的起始行 |
| startColumnIndex | int | 0 | list内容的起始列 |


使用方法：
```java
@ExcelSheet(name = "测试工作表", leftMargin = 0.1968504, rightMargin = 0.1968504, verticallyCenter = false)
@ExcelPrint(landscape = true, scale = 81)
public class TestEntity {

    @ExcelList(startRowIndex = 6)
    private List<TestListEntity> testListEntities;
}
```
#### @ExcelListColumn
结合@ExcelList注解进行使用。用于定义@ExcelList注解定义的List的T中的字段。

| 参数 | 类型 | 默认值 | 说明 |
| :----: | :----: | :----: | ---- |
| rowOffset | int | 0 | 起始行偏移量，即字段的内容显示的单元格开始行为@ExcelList中的startRowIndex+@ExcelListColumn中的rowOffset+@ExcelListTitle中的rowOffset+1 |
| columnOffset | int | 0 | 起始列偏移量，即字段的内容显示的单元格开始列为@ExcelList中的startRowIndex+@ExcelListColumn中的rowOffset |
| rowHeight | float | 13.5f | 行高 |
| nullValue | String |  | 当字段值为null时单元格的内容 |
| booleanTrueValue | String |  | 当字段值为boolean型且值为true时单元格的内容 |
| booleanFalseValue | String |  | 当字段值为boolean型且值为false时单元格的内容 |
| excelListTitle | @ExcelListTitle |  | 没有默认值，必须定义 |
| excelCellStyle | @ExcelCellStyle | @ExcelCellStyle  | 单元格样式，参见@ExcelCellStyle注解 |
| excelCellMerged | @ExcelCellMerged | @ExcelCellMerged  | 单元格合并方式，参见@ExcelCellMerged注解 |

使用方法：
```java
public class TestListEntity {

    @ExcelListColumn(
            rowHeight = 22,
            rowOffset = 1,
            columnOffset = 1,
            excelListTitle = @ExcelListTitle(
                    name = "序号",
                    rowHeight = 57,
                    columnWidth = 4.5f)
    )
    private String testListField;
}
```
#### @ExcelListTitle
定义@ExcelList注解定义的List的T中的字段内容对应的标题，该标题只体现一次，不重复体现。

| 参数 | 类型 | 默认值 | 说明 |
| :----: | :----: | :----: | ---- |
| name | String | 标题 | 标题名称 |
| rowOffset | int | 0 | 起始行偏移，不解释了 |
| columnOffset | int | 0 | 起始列偏移，不解释了 |
| rowHeight | float | 13.5f | 行高 |
| columnWidth | float | 8.38f | 列宽 |
| excelCellStyle | @ExcelCellStyle | @ExcelCellStyle  | 单元格样式，参见@ExcelCellStyle注解 |
| excelCellMerged | @ExcelCellMerged | @ExcelCellMerged  | 单元格合并方式，参见@ExcelCellMerged注解 |

使用方法：
```java
public class TestListEntity {

    @ExcelListColumn(
            rowHeight = 22,
            rowOffset = 1,
            columnOffset = 1,
            excelListTitle = @ExcelListTitle(
                    name = "序号",
                    rowHeight = 57,
                    columnWidth = 4.5f)
    )
    private String testListField;
}
```

#### @ExcelPicture
必须要有一个能插入图片的注解，请注意该注解只支持定义在byte[]类型的字段上。

| 参数 | 类型 | 默认值 | 说明 |
| :----: | :----: | :----: | ---- |
| col1 | int | 0 | 开始单元格列标 |
| row1 | int | 0 | 开始单元格行标 |
| col2 | int | 0 | 结束单元格列标 |
| row2 | int | 0 | 结束单元格行标 |
| dx1 | int | 0 | 开始单元格x坐标 |
| dy1 | int | 0 | 开始单元格y坐标 |
| dx2 | int | 0 | 结束单元格x坐标 |
| dy2 | int | 0 | 结束单元格y坐标 |

使用方法：
```java
@ExcelSheet(name = "测试工作表", leftMargin = 0.1968504, rightMargin = 0.1968504, verticallyCenter = false)
@ExcelPrint(landscape = true, scale = 81)
public class TestEntity {

    /**
     * 该图片则插入到单元格A1中且完全填充了整个A1单元格。懂的人自然懂，不懂的慢慢学
     */
    @ExcelPicture(row1 = 0, col1 = 0, row2 = 1, col2 = 1,dx1 = 0,dy1 = 0,dx2 = 0,dy2 = 0)
    private byte[] img;
}
```

----------

## 完整的使用例子

### 一个混合表格

定义一个混合的sheet类
```java
package com.toolsder.epoi;

import com.toolsder.epoi.annotation.field.*;
import com.toolsder.epoi.annotation.type.ExcelSheet;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;

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
    @ExcelCell(
            rowIndex = 0,
            columnIndex = 0,
            rowHeight = 20.0f,
            nullValue = "/",
            excelCellStyle = @ExcelCellStyle(
                    borderLeft = BorderStyle.THIN,
                    borderTop = BorderStyle.DASH_DOT_DOT,
                    borderRight = BorderStyle.HAIR,
                    borderBottom = BorderStyle.MEDIUM_DASH_DOT,
                    wrapText = true,
                    excelCellFont = @ExcelCellFont(
                            name = "微软雅黑",
                            heightInPoints = 20,
                            bold = true,
                            italic = true,
                            underline = Font.U_DOUBLE_ACCOUNTING
                    )
            ),
            excelCellMerged = @ExcelCellMerged(
                    rowCount = 2,
                    columnCount = 2
            )
    )
    private String title;

    /**
     * 一个非合并小标题
     */
    @ExcelCell(
            rowIndex = 2,
            columnIndex = 0,
            rowHeight = 21.0f,
            nullValue = "/",
            excelCellStyle = @ExcelCellStyle(
                    borderLeft = BorderStyle.THIN,
                    borderTop = BorderStyle.DASH_DOT_DOT,
                    borderRight = BorderStyle.HAIR,
                    borderBottom = BorderStyle.MEDIUM_DASH_DOT,
                    wrapText = true,
                    excelCellFont = @ExcelCellFont(
                            heightInPoints = 16,
                            bold = true,
                            italic = true,
                            underline = Font.U_DOUBLE_ACCOUNTING
                    )
            )
    )
    private String subTitle;

    /**
     * 定义列表内容，它从第三行的第二列开始
     */
    @ExcelList(
            startRowIndex = 3,
            startColumnIndex = 1
    )
    private List<MixingSheetList> mixingSheetLists;


    /**
     * 定义一个图片
     */
    @ExcelPicture(row1 = 0,col1 = 0,row2 = 1,col2 = 1)
    private byte[] img;

}

```
定义混合sheet类中的字段为list的T
```java
package com.toolsder.epoi;

import com.toolsder.epoi.annotation.field.*;
import com.toolsder.epoi.annotation.type.ExcelSheet;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;

/**
 * created by 猴子请来的逗比 On 2020/7/1
 * <p>
 * 混合的sheet中的list
 *
 * @author by 猴子请来的逗比
 */
@Getter
@Setter
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

```
生成sheet
```java
package com.toolsder.epoi;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * created by 猴子请来的逗比 On 2020/6/28
 *
 * @author by 猴子请来的逗比
 */
public class ExcelUtilsTest {
    @Test
    public void test() {
        //获取一个图片
        try (InputStream inputStream = new FileInputStream("D:\\1232.png")) {
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            BufferedImage bufferedImage = ImageIO.read(inputStream);
            ImageIO.write(bufferedImage, "PNG", byteArrayOutputStream);
            //定义一个XSSFWorkbook工作表
            XSSFWorkbook workbook = new XSSFWorkbook();
            //创建sheet表的类
            MixingSheetList mixingSheetList1 = new MixingSheetList();
            mixingSheetList1.setId("1");
            mixingSheetList1.setContent("内容1");
            MixingSheetList mixingSheetList2 = new MixingSheetList();
            mixingSheetList2.setId("2");
            mixingSheetList2.setContent("内容2");
            List<MixingSheetList> mixingSheetListList = new ArrayList<>();
            mixingSheetListList.add(mixingSheetList1);
            mixingSheetListList.add(mixingSheetList2);
            MixingSheet mixingSheet = new MixingSheet();
            mixingSheet.setTitle("标题1");
            mixingSheet.setSubTitle("标题2");
            mixingSheet.setMixingSheetLists(mixingSheetListList);
            mixingSheet.setImg(byteArrayOutputStream.toByteArray());
            //开始创建一个工作表对象
            ExcelSheetFile excelSheetFile = new ExcelSheetFile(mixingSheet, workbook);
            //向XSSFWorkbook中生成工作表
            excelSheetFile.generateSheet();
            //导出文件
            //捕获内存缓冲区的数据,转换成字节数组
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            try {
                workbook.write(os);
            } catch (Exception e) {
                e.printStackTrace();
            }
            try (FileOutputStream fileOutputStream = new FileOutputStream("D:\\test.xlsx")) {
                fileOutputStream.write(os.toByteArray());
            } catch (Exception e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

}

```

### 一个清单表格
定义一个清单sheet类
```java
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

```
生成sheet
```java
package com.toolsder.epoi;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * created by 猴子请来的逗比 On 2020/6/28
 *
 * @author by 猴子请来的逗比
 */
public class ExcelUtilsTest {
    @Test
    public void test() {
        //获取一个图片
        try (InputStream inputStream = new FileInputStream("D:\\1232.png")) {
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            BufferedImage bufferedImage = ImageIO.read(inputStream);
            ImageIO.write(bufferedImage, "PNG", byteArrayOutputStream);
            //定义一个XSSFWorkbook工作表
            XSSFWorkbook workbook = new XSSFWorkbook();
            //创建sheet表的类
            MixingSheetList mixingSheetList1 = new MixingSheetList();
            mixingSheetList1.setId("1");
            mixingSheetList1.setContent("内容1");
            MixingSheetList mixingSheetList2 = new MixingSheetList();
            mixingSheetList2.setId("2");
            mixingSheetList2.setContent("内容2");
            List<MixingSheetList> mixingSheetListList = new ArrayList<>();
            mixingSheetListList.add(mixingSheetList1);
            mixingSheetListList.add(mixingSheetList2);
            MixingSheet mixingSheet = new MixingSheet();
            mixingSheet.setTitle("标题1");
            mixingSheet.setSubTitle("标题2");
            mixingSheet.setMixingSheetLists(mixingSheetListList);
            mixingSheet.setImg(byteArrayOutputStream.toByteArray());
            //开始创建一个工作表对象
            ExcelSheetFile excelSheetFile = new ExcelSheetFile(mixingSheet, workbook);
            //向XSSFWorkbook中生成工作表
            excelSheetFile.generateSheet();
            //开始创建另一个工作表对象
            ExcelSheetFile excelSheetFile1 = new ExcelSheetFile(mixingSheetListList, workbook);
            //向XSSFWorkbook中生成工作表
            excelSheetFile1.generateSheet();

            //导出文件
            //捕获内存缓冲区的数据,转换成字节数组
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            try {
                workbook.write(os);
            } catch (Exception e) {
                e.printStackTrace();
            }
            try (FileOutputStream fileOutputStream = new FileOutputStream("D:\\test.xlsx")) {
                fileOutputStream.write(os.toByteArray());
            } catch (Exception e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

}

```

----------

## 结束语

真的结束了吗？当然没有了，还有Excel导入的功能还没有实现呢，现在正在构思导入的功能，不久之后将会与大家见面！！！

当然了，作为一个小小小小小小的开源爱好者，在获得大家的精神支持的时候也还需要一点点点点点点物质支持。
**支付宝**：<img src="https://github.com/smallQ1017/pay-qrcode/blob/master/zhifubao-qrcode.png?raw=true" width="150" height="150"/>**微信**：<img src="https://github.com/smallQ1017/pay-qrcode/blob/master/weixin-qrcode.png?raw=true" width="150" height="150"/>**PayPal**：https://paypal.me/gentlyding?locale.x=zh_XC

再次感谢大家！！！
