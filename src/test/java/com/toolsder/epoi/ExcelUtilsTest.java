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
