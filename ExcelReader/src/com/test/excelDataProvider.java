package com.test;

//import java.util.ArrayList;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.testng.annotations.*;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class excelDataProvider {

	myExtentReports htmlReport= new myExtentReports();
	
	//static ExtentTest test;
	//static ExtentReports report;

	/*ActLogin actLogin = new ActLogin();

    @BeforeTest
	@Parameters({"url"})
    public void beforeTest(String url) throws Exception {
		System.out.println("@beforeTest URL is: "+url);
		//actLogin.fbBrowserLaunch();
    }
	
	@BeforeClass
    @Parameters({"username", "password"})
	public void beforeClass(String username, String password) throws Exception {
		System.out.println("@BeforeClass username is: "+username+" and Password is: "+password);
    }
*/
	
	
    @BeforeMethod
    public void beforeMethod() throws Exception {
    }

    @AfterMethod
    public void afterMethod() throws Exception {
    }

    @DataProvider
	public Iterator<Object[]> getTestData(){
		int s= 1, e=4;
		ArrayList<Object[]> testData=FullExcelReader.xlSheetReader(".\\testData\\DataSheet.xlsx");

		System.out.println(	testData.size());
		
		
//		testData.remove(0);

		for (int i=testData.size()-1; i>e;i--) {
			testData.remove(i);
			System.out.println(	testData.size());
		}
		
		for (int i=0; i<s;i++) {
			testData.remove(0);
			System.out.println(	testData.size());
		}


		
		
		
		return testData.iterator();
	}

	@Test(dataProvider="getTestData")
	public void NormalEstabEntryPermitTourism(String TestCaseName, String ServiceName) throws Exception {		
		//ArrayList<Object[]> testData=FullExcelReader.xlSheetReader(".\\testData\\DataSheet.xlsx");
		//ArrayList<Object> mydate = new ArrayList<Object>();
		if(ServiceName.contains("TestCase3")) {
			System.out.println("This is test case 3 testing");			
		}
		
		System.out.println(TestCaseName);
		System.out.println(ServiceName);
		htmlReport.test.log(LogStatus.PASS, TestCaseName);
	}

/*If you want to execute specific rows
 * 
 * 	@Test
	public void NormalEstabEntryPermitTourism() throws Exception {
		String TestCaseName;
		String ServiceName;

		List<Object> row = FullExcelReader.xlRowReader(".\\testData\\DataSheet.xlsx",1);
		TestCaseName= row.get(0).toString();
		ServiceName= row.get(1).toString();
	    
		if(ServiceName.contains("TestCase3")) {
			System.out.println("This is test case 3 testing");			
		}
		
		System.out.println(TestCaseName);
		System.out.println(ServiceName);
		htmlReport.test.log(LogStatus.PASS, TestCaseName);
	}

    
	@Parameters({"startTestCaseRow","endTestCaseRow"})    
	@Test
	public void NormalEstabEntryPermitTourism1(int startTestCaseRow, int endTestCaseRow) throws Exception {
		String TestCaseName;
		String ServiceName;
		System.out.println(startTestCaseRow);
		System.out.println(endTestCaseRow);

		for(int i=startTestCaseRow; i<= endTestCaseRow;i++) {
			List<Object> row = FullExcelReader.xlRowReader(".\\testData\\DataSheet.xlsx",i);
			TestCaseName= row.get(Constants.TestCaseName).toString();
			ServiceName= row.get(Constants.ServiceName).toString();
			
			if(ServiceName.contains("TestCase3")) {
				System.out.println("This is test case 3 testing");			
			}
			System.out.println(TestCaseName);
			System.out.println(ServiceName);
			htmlReport.test.log(LogStatus.FAIL, TestCaseName);
		}	
	}
	
*/	
    
    
	@BeforeClass
	public void beforeClass() throws Exception {
		htmlReport.startTest();
    }

	
    @AfterClass
    public void afterClass() throws Exception {
    	htmlReport.endTest();
    }

}
