package com.lw.tech;

public class MainExcelFile {
    public static void main(String[] args){
        String excelpath="C:\\Users\\LENOVO\\Documents\\databaseapp\\ExcelFile\\StudentInformation.xlsx";
        String sheetName="Sheet1";
        Excelutils excel=new Excelutils(excelpath,sheetName);
        excel.getRowCount();
        excel.getCellData(1,0);
        excel.getCellData(1,1);
        excel.getCellData(1,2);

    }
}
