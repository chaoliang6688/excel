package com.lc;



import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author Alike
 * @create 2020 05 18 11:51
 */
public class ExcelReadTest {

    String PATH = "E:\\IdeaCode\\excel-poi\\";

    @Test
    public void testRead03() throws IOException {
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "03.xls");
        //创建一个HSSFWorkbook对象
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        //获取工作表getSheet(表名)/getSheetAt(0):根据工作表的下标
        Sheet sheet = workbook.getSheetAt(0);
        //获取行
        Row row = sheet.getRow(0);
        //获取单元格/列
        Cell cell = row.getCell(0);
        //获取值的时候，注意内容的类型
        //获取单元格的内容
        System.out.println(cell.getStringCellValue());
        fileInputStream.close();


    }

    @Test
    public void testRead07() throws IOException {
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "07.xlsx");
        //创建一个HSSFWorkbook对象
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        //获取工作表getSheet(表名)/getSheetAt(0):根据工作表的下标
        Sheet sheet = workbook.getSheetAt(0);
        //获取行
        Row row = sheet.getRow(0);
        //获取单元格/列
        Cell cell = row.getCell(0);
        //获取值的时候，注意内容的类型
        //获取单元格的内容
        System.out.println(cell.getStringCellValue());
        fileInputStream.close();


    }

    @Test
    public void testCellType() throws IOException {
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "07.xlsx");
        //创建一个HSSFWorkbook对象
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        //获取工作表getSheet(表名)/getSheetAt(0):根据工作表的下标
        Sheet sheet = workbook.getSheetAt(0);
        //获取标题内容
        Row rowTitle = sheet.getRow(0);
        if (rowTitle != null){
            //获取标题总数
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0;cellNum < cellCount;cellNum++){
                //获取单元格
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null){
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue+ "|");
                }
            }
            System.out.println();
        }
        //获取表中的内容
        //获取所有行数量
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1;rowNum < rowCount;rowNum++){
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null){
                //读取列的数据
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount;cellNum++){
                    Cell cell = rowData.getCell(cellNum);
                    //匹配列的数据类型
                    if(cell != null){
                        int cellType = cell.getCellType();
                        String cellValue = "";
                        switch (cellType){
                            case XSSFCell.CELL_TYPE_STRING://字符串
                                cellValue = cell.getStringCellValue();
                                break;
                            case XSSFCell.CELL_TYPE_BOOLEAN://布尔值
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case XSSFCell.CELL_TYPE_BLANK://空
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC://数字  (日期、普通数字)
                                if (HSSFDateUtil.isCellDateFormatted(cell)){//日期
                                    cellValue = new DateTime(cell.getDateCellValue()).toString("yyyy-MM-dd");
                                }else{
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;

                            case XSSFCell.CELL_TYPE_ERROR://错误
                                break;

                        }
                        System.out.print(cellValue+" | ");
                    }

                }
                System.out.println();
            }
        }
        fileInputStream.close();
    }
    @Test
    public void testFormula() throws IOException {
        FileInputStream fileInputStream = new FileInputStream("公式.xlsx");
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(5);
        Cell cell = row.getCell(0);
        //获取计算公式
       FormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
        //输出单元格的内容
        int cellType = cell.getCellType();
        switch (cellType){
            case Cell.CELL_TYPE_FORMULA://公式
                String formula = cell.getCellFormula();
                //计算
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue);
                break;
        }

    }
}
