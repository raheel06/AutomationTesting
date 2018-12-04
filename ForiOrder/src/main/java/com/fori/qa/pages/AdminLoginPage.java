package com.fori.qa.pages;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.fori.qa.base.TestBase;

public class AdminLoginPage extends TestBase{
	
	//Page Factory - OR:
	
	@FindBy(xpath="//*[@id=\"login\"]/p[1]")
	WebElement msgwelcome;
	
	@FindBy(id="loginform")
	WebElement formLogin;
	
	@FindBy(id="user_login")
	WebElement username;
	
	@FindBy(id="user_pass")
	WebElement password;
	
	@FindBy(id="wp-submit")
	WebElement loginBtn;
	
	@FindBy(xpath="//*[@id=\"login\"]/h1/a")
	WebElement foriLogo;
	
	@FindBy(xpath="//*[@id=\"wp-admin-bar-my-account\"]/a")
	WebElement account;
	
	@FindBy(xpath="//*[@id=\"wp-admin-bar-logout\"]/a")
	WebElement logout;
	
	@FindBy(xpath="//*[@id=\"login\"]/p[1]")
	WebElement msgLogout;
	
	
	//Initializing the Page Objects:
	public AdminLoginPage(){
		PageFactory.initElements(driver, this);
	}
	
	//Actions:
	public String validateLoginPageTitle(){
		return driver.getTitle();
	}
	
	public boolean validateLogoImage(){
		return foriLogo.isDisplayed();
	}
	
	public HomePage login(String un, String pwd){
		if(msgwelcome.isDisplayed()) {
			username.sendKeys(un);
			password.sendKeys(pwd);
			loginBtn.click();
		}
		//return account.isDisplayed();
		return new HomePage();
	}
	
	public String logout(){
		Actions mouse= new Actions(driver);
		mouse.moveToElement(account).perform();
		mouse.moveToElement(logout).click().perform();
		return msgLogout.getText();
	}
	
	
	
}
