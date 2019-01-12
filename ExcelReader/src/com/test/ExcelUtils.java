/**
 *  @author Juber ******************************************
 */

package com.emaratech.utils;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelUtils {
	private static XSSFSheet ExcelWSheet;
	private static XSSFWorkbook ExcelWBook;
	private static XSSFCell Cell;
	private static XSSFRow Row;



	//This method is to set the File path and to open the Excel file, Pass Excel Path and Sheetname as Arguments to this method
	public static void setExcelFile(String Path,String SheetName) throws Exception {
		try {
			// Open the Excel file
			FileInputStream ExcelFile = new FileInputStream(Path);
			// Access the required test data sheet
			ExcelWBook = new XSSFWorkbook(ExcelFile);
			ExcelWSheet = ExcelWBook.getSheet(SheetName);
			System.out.println(ExcelWSheet);
			//Log.info("Excel sheet opened");
		} catch (Exception e){
			throw (e);
		}
	}


	//This method reads the test data from the Excel cell.
	//We are passing row number and column number as parameters.
	public static String getCellData(int RowNum, int ColNum) {
		try {
			Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
			DataFormatter formatter = new DataFormatter();
			String cellData = formatter.formatCellValue(Cell);
			return cellData;
		} catch (Exception e) {
			throw (e);
		}
	}

	//This method is to write in the Excel cell, Row num and Col num are the parameters
	@SuppressWarnings("static-access")
	public static void setCellData(String Result,  int RowNum, int ColNum,String FileName) throws Exception    {
		try{


			Row  = ExcelWSheet.getRow(RowNum);
			Cell = Row.getCell(ColNum);
			if (Cell == null) {
				Cell = Row.createCell(ColNum);
				Cell.setCellValue(Result);
			} else {
				Cell.setCellValue(Result);
			}
			// Constant variables Test Data path and Test Data file name
			FileOutputStream fileOut = new FileOutputStream(Constant_EP.Path_TestData + FileName);
			ExcelWBook.write(fileOut);
			// fileOut.flush();
			fileOut.close();

			// Reload the workbook, workaround for bug 49940
			// https://issues.apache.org/bugzilla/show_bug.cgi?id=49940
			//ExcelWBook = new XSSFWorkbook(new FileInputStream(Constant.Path_TestData + Constant.File_TestData));
		}catch(Exception e){
			throw (e);
		}
	}

	public static int getRowContains(String sTestCaseName, int colNum) throws Exception{
		int i;
		try {
			int rowCount = ExcelUtils.getRowUsed();
			for ( i=0 ; i<rowCount; i++){
				if  (ExcelUtils.getCellData(i,colNum).equalsIgnoreCase(sTestCaseName)){
					break;
				}
			}
			System.out.println(sTestCaseName+ "::"+ (i));
			return i;
		}catch (Exception e){
			//Log.error("Class ExcelUtil | Method getRowContains | Exception desc : " + e.getMessage());
			throw(e);
		}
	}

	public static int getRowUsed() throws Exception {
		try{
			int RowCount = ExcelWSheet.getLastRowNum();
			//	Log.info("Total number of Row used return as < " + RowCount + " >.");		
			return RowCount;
		}catch (Exception e){
			//	Log.error("Class ExcelUtil | Method getRowUsed | Exception desc : "+e.getMessage());
			System.out.println(e.getMessage());
			throw (e);
		}

	}
	public static void writeWorkbook(String FileName)throws Exception{

		FileOutputStream fileOut = new FileOutputStream(Constant_EP.Path_TestData + FileName);
		ExcelWBook.write(fileOut);
		// fileOut.flush();
		fileOut.close();
	}

	public static void SetColorFailed(String FileName)throws Exception{

		FileOutputStream fileOut = new FileOutputStream(Constant_EP.Path_TestData + FileName);
		ExcelWBook.write(fileOut);
		// fileOut.flush();
		fileOut.close();
	}

	public static void clearCells()throws Exception{

		//Clear previous result 

	}

	public static Iterator<org.apache.poi.ss.usermodel.Row> getCurrentSheetRows() {
		return ExcelWSheet.iterator();
	}

	public static int getTestLastRow(String TestCaseName,int iTestCaseRow)throws Exception{

		int i=0;
		System.out.println(getRowUsed());
		try{

			for(i=iTestCaseRow+1;i<=getRowUsed();i++){

				String testCaseName=ExcelUtils.getCellData(i,2);

				if("".equals(testCaseName)){

				}else{

					break;

				}

			}

			//retrun last row
			return i;        

		}catch(Exception e){

			System.out.println(e.getMessage());

			throw(e);

		}

	}


}