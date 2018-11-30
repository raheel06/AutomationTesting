package com.Extra;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class excelutil {
	public static void setExcelFile() throws Exception {
		XSSFSheet ExcelWSheet;
		XSSFWorkbook ExcelWBook;
		XSSFCell Cell;
		XSSFRow Row;

		// Open the Excel file
		FileInputStream ExcelFile = new FileInputStream(".\\testData\\DataSheet.xlsx");
		// Access the required test data sheet
		ExcelWBook = new XSSFWorkbook(ExcelFile);
		ExcelWSheet = ExcelWBook.getSheet("TestData");
		System.out.println(ExcelWSheet);
		//Log.info("Excel sheet opened");

	}		
}
