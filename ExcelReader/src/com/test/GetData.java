
import java.util.ArrayList;
import java.util.List;
import com.emaratech.utils.ExcelReader;
import java.nio.file.Path;
import java.nio.file.Paths;

public class GetData {
	public static String TestCaseName;
	public static String ServiceName;
	
	public static String Col_TestCaseName ;
	public static String Col_ServiceName ;
	
	static Path currentRelativePath = Paths.get("");
	static String rootPath = currentRelativePath.toAbsolutePath().toString();

	//private static String rootPath = UtilityClass.GetAppSetttingsByKey("s"); 

	public static String Path_TestData = rootPath + "\\TestData\\";

	public static String Path_ScreenShot = rootPath + "\\src\\main\\java\\";
	
	public static String File_TestData_Transit = "Transit.xlsx";
	public static String File_TestData_Airline = "Airline.xlsx";
	public static String File_TestData_Tourism = "Tourism.xlsx";
	public static String File_TestData_GCCEP = "GCCEP.xlsx";
	
	public void fnInputTestData(String TestCaseNameInTestDataFile) {
		
		//List<Object> row = ExcelReader.xlRowReader(".\\TestData\\DataSheet.xlsx", rowNum);
		List<Object> row = ExcelReader.xlRowReaderWithName(".\\TestData\\DataSheet.xlsx", TestCaseNameInTestDataFile);
		//		TestCaseName= row.get(0).toString();
		//		ServiceName= row.get(1).toString();

		TestCaseName=              	row.get(0).toString();
		ServiceName=               	row.get(1).toString();
	}

	
	public static void fnTestDataRowWise(int iTestCaseRow){ 
		//public static final String Payment_Type = "";
		//public static final String Applications_Options = "";
		
		//Test Data Sheet Columns
		Col_TestCaseName=			ExcelUtils.getCellData(iTestCaseRow,0);
		Col_ServiceName=			ExcelUtils.getCellData(iTestCaseRow,1);
	}
	
	
public static List<Integer> fnGetTestSet(String TestCaseName, String TestFileName) throws Exception {
	int iTestCaseStart;
	int iTestCaseLastEnd;
	List<Integer> items = new ArrayList<Integer>();
	
	ExcelUtils.setExcelFile(Constant_EP.Path_TestData+TestFileName,"TestData");
	iTestCaseStart=ExcelUtils.getRowContains(TestCaseName,Constant_EP.Col_TestCaseName);
	iTestCaseLastEnd=ExcelUtils.getTestLastRow(TestCaseName,iTestCaseStart);
	
	items.add(iTestCaseStart);
	items.add(iTestCaseLastEnd);
	
	return items;
	
//	iTestCaseRow=ExcelUtils.getRowContains(method.getName(),Constant_EP.Col_TestCaseName);
//	iTestCaseLastRow=ExcelUtils.getTestLastRow(method.getName(),iTestCaseRow);
	
}


}
