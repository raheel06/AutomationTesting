package com.fori.qa.tests;

import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import com.fori.qa.base.TestBase;
import com.fori.qa.pages.HomePage;
import com.fori.qa.pages.AdminLoginPage;

public class AdminLoginPageTest extends TestBase{
	AdminLoginPage loginPage;
	HomePage homePage;
	
	public AdminLoginPageTest(){
		super();
	}
	
	@BeforeMethod
	public void setUp(){
		
	}
	
	@Test
	public void loginTest(){
		initialization();
		loginPage = new AdminLoginPage();
		
		driver.get(prop.getProperty("adminURL"));
		
		
		String title = loginPage.validateLoginPageTitle();
		Assert.assertEquals(title.contains("ForiOrder.pk"), true);

		boolean flag = loginPage.validateLogoImage();
		Assert.assertTrue(flag);

		homePage = loginPage.login(prop.getProperty("username"), prop.getProperty("password"));
		//String dashTitle = loginPage.validateLoginPageTitle();
		//Assert.assertEquals(dashTitle.contains("Dashboard"), true);
		
		//String logoutMsg = loginPage.logout();
		//Assert.assertEquals(logoutMsg.contains("You are now logged out"), true);
		driver.quit();
	}
	
	@Test
	public void LogoImageTest(){
		initialization();
		loginPage = new AdminLoginPage();
		
		driver.get(prop.getProperty("adminURL"));
		
		String title = loginPage.validateLoginPageTitle();
		Assert.assertEquals(title.contains("ForiOrder.pk"), true);

		boolean flag = loginPage.validateLogoImage();
		Assert.assertTrue(flag);

		homePage = loginPage.login(prop.getProperty("username"), prop.getProperty("password"));
		driver.quit();
	}
//	
//	@Test(priority=3)
//	public void loginTest(){
//	}
	
	@AfterMethod
	public void tearDown(){
		
	}
	
	
	
	

}
