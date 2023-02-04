package com.lw.tech;


import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class Excelutils {
    static XSSFWorkbook workbook;
    static XSSFSheet sheet;

    public Excelutils(String excelpath,String sheetName){
        try {
            workbook = new XSSFWorkbook(excelpath);
            sheet = workbook.getSheet(sheetName);
        }catch(Exception e){
            e.printStackTrace();
        }
    }

      public static void getCellData(int rowNum, int colNum){
          try {
              /*
              String exclePath = "C:\\Users\\LENOVO\\Documents\\databaseapp\\ExcelFile\\StudentInformation.xlsx";
              XSSFWorkbook workbook = new XSSFWorkbook(exclePath);
              XSSFSheet sheet = workbook.getSheet("Sheet1");
              //Object value = sheet.getRow(1).getCell(1).getStringCellValue();
               */
              DataFormatter formatter=new DataFormatter();
              Object value=formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
              System.out.println(value);
          }
          catch(Exception e){
              e.printStackTrace();
          }
      }
      public static void getRowCount(){
          try {
              /*
              String exclePath = "C:\\Users\\LENOVO\\Documents\\databaseapp\\ExcelFile\\StudentInformation.xlsx";
              XSSFWorkbook workbook = new XSSFWorkbook(exclePath);
              XSSFSheet sheet = workbook.getSheet("sheet1");

               */
              int rowcount = sheet.getPhysicalNumberOfRows();
              System.out.println("No of Rows: " + rowcount);
          }catch(Exception e){
              e.printStackTrace();
          }
      }

}