package com.test;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.function.ToIntFunction;

import org.apache.poi.hpsf.Array;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class FullExcelReader {

	public static ArrayList<Object[]> xlSheetReader(String xlPath) {
		final String SAMPLE_XLSX_FILE_PATH = xlPath;
		ArrayList<Object[]> tableData = new ArrayList<Object[]>();
				
		try {
			Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));			
			Sheet sheet = workbook.getSheetAt(0);
			// Create a DataFormatter to format and get each cell's value as String
			DataFormatter dataFormatter = new DataFormatter();

			// 1. Or you can use a for-each loop to iterate over the rows and columns
	        //System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
	        for (Row row: sheet) {
	        	List<Object> rowData = new ArrayList<Object>();
	            for(Cell cell: row) {
	                String cellValue = dataFormatter.formatCellValue(cell);
	                //System.out.print(cellValue + "\t");
	                rowData.add(cellValue);
	            }
	            tableData.add(rowData.toArray());
	            //System.out.println();
	        }
            	        
		    /*//2. You can obtain a rowIterator and columnIterator and iterate over them
	        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
	        Iterator<Row> rowIterator = sheet.rowIterator();
	        while (rowIterator.hasNext()) {
	        	Row row = rowIterator.next();
	            // Now let's iterate over the columns of the current row
	            Iterator<Cell> cellIterator = row.cellIterator();
	            while (cellIterator.hasNext()) {
	            	Cell cell = cellIterator.next();
	            	String cellValue = dataFormatter.formatCellValue(cell);
	            	System.out.print(cellValue + "\t");	
	            }
	            System.out.println();
	        }*/
	        
	     /*// Retrieving the number of sheets in the Workbook
			System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
			System.out.println("Sheets Names:");
			for(Sheet xlSheet: workbook) {
				System.out.println("=> " + xlSheet.getSheetName());
			}*/
	        return tableData;
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		return null;
	}
	
	public static List<Object> xlRowReader(String xlPath, int rowNum) {
		final String SAMPLE_XLSX_FILE_PATH = xlPath;
		
		try {
			List<Object> rowData = new ArrayList<Object>();
            
			Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));			
			Sheet sheet = workbook.getSheetAt(0);
			DataFormatter dataFormatter = new DataFormatter();
			Row row= sheet.getRow(rowNum);
			
        	for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                //System.out.print(cellValue + "\t");
                rowData.add(cellValue);
            }
        	return rowData;
		}
		catch (Exception e) {
			System.out.println(e.getMessage());
		}
		
		return null;
	
	}
}