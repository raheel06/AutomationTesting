package com.test;

import org.testng.annotations.Test;

public class ActLogin {
	String UserType="";

	public void fnLogin(String usertype) {
		System.out.println("Login Successfull");
	}
	
	public void fbBrowserLaunch() {
		try {
			System.out.println("Browser Launched: fbBrowserLaunch()");
		}
		catch (AssertionError ex){
			throw ex;
		}
	}
	
}
