package com.test;

import java.util.ArrayList;
import java.util.List;
import java.nio.file.Path;
import java.nio.file.Paths;

public class GetData {
	public static int TestCase_Start=1;
	public static int TestCase_End=2;

	public static String TestCaseName;
	public static String ServiceName;
	
	public static int ind_TestCaseName=0;
	public static int ind_ServiceName=1;
	
	public static String Col_TestCaseName ;
	public static String Col_ServiceName ;
	
	static Path currentRelativePath = Paths.get("");
	static String rootPath = currentRelativePath.toAbsolutePath().toString();
	//private static String rootPath = UtilityClass.GetAppSetttingsByKey("s"); 
	public static String Path_TestData = rootPath + "\\TestData\\";
	public static String Path_ScreenShot = rootPath + "\\src\\main\\java\\";
	
	public static String File_TestData = "TestData.xlsx";

	
	public void fnInputTestData(String TestCaseNameInTestDataFile) {
		//List<Object> row = ExcelReader.xlRowReader(".\\TestData\\DataSheet.xlsx", rowNum);
		List<Object> row = ExcelReader.xlRowReaderWithName(".\\TestData\\TestData.xlsx", TestCaseNameInTestDataFile);
		TestCaseName=              	row.get(0).toString();
		ServiceName=               	row.get(1).toString();
	}

	
	public static void fnTestDataRowWise(int iTestCaseRow) throws Exception{ 
		//public static final String Payment_Type = "";
		//public static final String Applications_Options = "";
		
//		ExcelUtils.setExcelFile(Path_TestData+TestFileName, "TestData");
		
		//Test Data Sheet Columns
		Col_TestCaseName=			ExcelUtils.getCellData(iTestCaseRow,0);
		Col_ServiceName=			ExcelUtils.getCellData(iTestCaseRow,1);
	}
	
	
	
	public static List<Integer> fnGetTestSet(String TestCaseName, String TestFileName) throws Exception {
		int iTestCaseStart;
		int iTestCaseLastEnd;
		List<Integer> items = new ArrayList<Integer>();
		
		ExcelUtils.setExcelFile(Path_TestData+TestFileName, "TestData");
		iTestCaseStart=ExcelUtils.getRowContains(TestCaseName, ind_TestCaseName);
		iTestCaseLastEnd=ExcelUtils.getTestLastRow(TestCaseName, iTestCaseStart);
		
		items.add(iTestCaseStart);
		items.add(iTestCaseLastEnd);
		
		return items;
		
	//	iTestCaseRow=ExcelUtils.getRowContains(method.getName(),Constant_EP.Col_TestCaseName);
	//	iTestCaseLastRow=ExcelUtils.getTestLastRow(method.getName(),iTestCaseRow);
		
	}

	public static void fnSetTestSheet(String TestFileName) throws Exception {	
		ExcelUtils.setExcelFile(Path_TestData+TestFileName, "TestData");		
	}

}
