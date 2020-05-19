package com.lc;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author Alike
 * @create 2020 05 18 10:19
 */
public class ExcelWirteTest {
    String PATH = "E:\\IdeaCode\\excel-poi\\";

    @Test
    public void testWrite03() throws IOException {
        //1.创建工作薄
        Workbook workbook = new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("全球油价统计");
        //3.创建一行(1,1)
        Row row1 = sheet.createRow(0);
        //4.创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("国际行情");

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("国外行情");

        //创建第二行(2,1)
        Row row2 = sheet.createRow(1);
        //(2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        //(2,2)
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-mm-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表(io流)   03版就是使用xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "测试表1.xls");
        //输出
        workbook.write(fileOutputStream);

        //关闭流
        fileOutputStream.close();
        System.out.println("文件生成完毕");

    }
    @Test
    public void testWrite07() throws IOException {
        //1.创建工作薄
        Workbook workbook = new XSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("全球油价统计");
        //3.创建一行(1,1)
        Row row1 = sheet.createRow(0);
        //4.创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("国际行情");

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("国外行情");

        //创建第二行(2,1)
        Row row2 = sheet.createRow(1);
        //(2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        //(2,2)
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-mm-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表(io流)   03版就是使用xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "测试表2.xlsx");
        //输出
        workbook.write(fileOutputStream);

        //关闭流
        fileOutputStream.close();
        System.out.println("文件生成完毕");

    }

    @Test
    public void testWrite03BigData() throws IOException {
        //时间开始
        long begin = System.currentTimeMillis();
        //创建工作薄
        Workbook workbook = new HSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet("测试大数据量");
        //写入数据
        for (int rowNum = 0;rowNum < 65536;rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0;cellNum < 10;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "测试03版大数据量.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        //时间结束
        long end = System.currentTimeMillis();

        System.out.println("整个过程消耗的时间:"+(double) (end-begin) / 1000);//3.703
    }
    @Test
    public void testWrite07BigData() throws IOException {
        //时间开始
        long begin = System.currentTimeMillis();
        //创建工作薄
        Workbook workbook = new XSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet("测试大数据量");
        //写入数据
        for (int rowNum = 0;rowNum < 65536;rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0;cellNum < 10;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "测试07版大数据量.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        //时间结束
        long end = System.currentTimeMillis();

        System.out.println("整个过程消耗的时间:"+(double) (end-begin) / 1000);//14.862
    }

    @Test
    public void testWrite07BigDataS() throws IOException {
        //时间开始
        long begin = System.currentTimeMillis();
        //创建工作薄
        Workbook workbook = new SXSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet("测试大数据量");
        //写入数据
        for (int rowNum = 0;rowNum < 65536;rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0;cellNum < 10;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "测试07版升级版SXSSFWorkbook大数据量.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        //清除临时文件
        ((SXSSFWorkbook)workbook).dispose();
        //时间结束
        long end = System.currentTimeMillis();

        System.out.println("整个过程消耗的时间:"+(double) (end-begin) / 1000);//3.581
    }

}

