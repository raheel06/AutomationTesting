package com.test;

import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class myExtentReports {
	static ExtentTest test;
	static ExtentReports report;

	@BeforeClass
	public static void startTest()
	{
		//report = new ExtentReports(System.getProperty("user.dir")+"\\ExtentReportResults.html");
		report = new ExtentReports("D:\\AutomationTesting\\ExcelReader\\ExtentReportResults.html");
		test = report.startTest("NormalEstabEntryPermitTourism");
	}
	
	//@Test
	public void extentReportsDemo()
	{
		test.log(LogStatus.PASS, "Navigated to the specified URL");
	}
	
	//@AfterClass
	public static void endTest()
	{
		report.endTest(test);
		report.flush();
	}
}