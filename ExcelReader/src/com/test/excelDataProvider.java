package com.test;

//import java.util.ArrayList;
import java.util.ArrayList;
import java.util.Iterator;
import org.testng.annotations.*;

public class excelDataProvider {
	
	@DataProvider
	public Iterator<Object[]> getTestData(){
		ArrayList<Object[]> testData=FullExcelReader.xlReader(".\\testData\\DataSheet.xlsx");
		testData.remove(0);
		return testData.iterator();
	}
	
	@Test(dataProvider="getTestData")
	public void excel(String TestCaseName, String ServiceName) throws Exception {
		// TODO Auto-generated method stub
		FullExcelReader.xlReader(".\\testData\\DataSheet.xlsx");
		//excelutil.setExcelFile();
		//ArrayList<Object> mydate = new ArrayList<Object>();
		System.out.println(TestCaseName);
		System.out.println(ServiceName);
	}
	
}
