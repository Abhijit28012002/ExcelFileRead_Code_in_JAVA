package com.lw.tech;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;

public class MainExcelFile {
    public static void main(String[] args){
        String excelpath="C:\\Users\\LENOVO\\Documents\\databaseappV2\\ExcelFile\\StudentInformation.xlsx";
        try {
            FileInputStream inputstream=new FileInputStream(excelpath);
            XSSFWorkbook workbook= new XSSFWorkbook(inputstream);
            XSSFSheet sheet= workbook.getSheetAt(0);
                                     ///iterator///
            Iterator iterator= sheet.iterator();
            while(iterator.hasNext()){
                XSSFRow row= (XSSFRow) iterator.next();
                Iterator cellIterator =row.cellIterator();
                while(cellIterator.hasNext()){
                    XSSFCell cell= (XSSFCell) cellIterator.next();
                    switch (cell.getCellType()){
                        case STRING : System.out.print(cell.getStringCellValue()); break;
                        case NUMERIC:  System.out.print(cell.getNumericCellValue()); break;
                        case BOOLEAN:  System.out.print(cell.getBooleanCellValue()); break;
                    }
                    System.out.print("     ");
                }
                System.out.println();
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        

    }
}
