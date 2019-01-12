package com.emaratech.utils;

import java.util.ArrayList;
import java.util.List;
import com.emaratech.utils.ExcelReader;
import java.nio.file.Path;
import java.nio.file.Paths;

public class GetData {
	public static String permitNo;
	public static String resNo;
	
	
	public static String SrNo;
	public static String TCJiraID;
	public static String TestCaseName;
	public static String ServiceName;
	public static String EstablishmentType;
	public static String ApplicantType;
	public static String PersonType;
	public static String VisitPurpose;;
	public static String Duration;;
	public static String Entry;
	public static String PassportType;
	public static String CurrentNationality;
	public static String PreviousNationality;
	public static String PassportIssueCountry;
	public static String ResidenceNumberOfYears;
	public static String BirthCountry;
	public static String Gender;
	public static String MaritalStatus;
	public static String RelationWithSponsor;
	public static String Religion;
	public static String Faith;
	public static String Education;
	public static String Profession;
	public static String FirstLanguage;
	public static String Emirate;
	public static String City;
	public static String Area;
	public static String CountryOutsideUAE;
	public static String TestCasePrereq;
	public static String GCCIssueCountry;
	public static String GCCPortOfArrival;
	public static String GCCArrivingFromCountry;
	public static String GCCArrivingFrom;
	public static String GCCSponsorType;
	public static String GCCSponsorNationality;
	public static String TransitComingFrom;
	public static String TransitLeavingTo;
	
	public static String Col_SrNo ;
	public static String Col_TCJiraID ;
	public static String Col_TestCaseName ;
	public static String Col_ServiceName ;
	public static String Col_EstablishmentType ;
	public static String Col_ApplicantType ;
	public static String Col_PersonType ;
	public static String Col_VisitPurpose ;
	public static String Col_Duration;
	public static String Col_Entry ;
	public static String Col_PassportType ;
	public static String Col_CurrentNationality ;
	public static String Col_PreviousNationality ;
	public static String Col_PassportIssueCountry ;
	public static String Col_ResidenceNumberOfYears ;
	public static String Col_BirthCountry ;
	public static String Col_Gender;
	public static String Col_MaritalStatus;
	public static String Col_RelationWithSponsor;
	public static String Col_Religion;
	public static String Col_Faith;
	public static String Col_Education;
	public static String Col_Profession;
	public static String Col_FirstLanguage;
	public static String Col_Emirate;
	public static String Col_City;
	public static String Col_Area;
	public static String Col_CountryOutsideUAE;
	public static String Col_TestCasePrereq ;
	public static String Col_GCCIssueCountry ;
	public static String Col_GCCPortOfArrival;
	public static String Col_GCCArrivingFromCountry ;
	public static String Col_GCCArrivingFrom ;
	public static String Col_GCCSponsorType ;
	public static String Col_GCCSponsorNationality ;
	public static String Col_TransitComingFrom ;
	public static String Col_TransitLeavingTo ;
	public static String Col_Result ;
	
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

		SrNo=						row.get(0).toString();
		TCJiraID=                  	row.get(1).toString();
		TestCaseName=              	row.get(2).toString();
		ServiceName=               	row.get(3).toString();
		EstablishmentType=          row.get(4).toString();
		ApplicantType=             	row.get(5).toString();
		PersonType=					row.get(6).toString();
		VisitPurpose=             	row.get(7).toString();
		Duration=                 	row.get(8).toString();
		Entry=                     	row.get(9).toString();
		PassportType=              	row.get(10).toString();
		CurrentNationality=        	row.get(11).toString();
		PreviousNationality=       	row.get(12).toString();
		PassportIssueCountry=		row.get(13).toString();
		ResidenceNumberOfYears=		row.get(14).toString();
		BirthCountry=              	row.get(15).toString();
		Gender=                    	row.get(16).toString();
		MaritalStatus=             	row.get(17).toString();
		RelationWithSponsor=       	row.get(18).toString();
		Religion=                  	row.get(19).toString();
		Faith=                     	row.get(20).toString();
		Education=                 	row.get(21).toString();
		Profession=                	row.get(22).toString();
		FirstLanguage=             	row.get(23).toString();
		Emirate=                   	row.get(24).toString();
		City=                      	row.get(25).toString();
		Area=                      	row.get(26).toString();
		CountryOutsideUAE=         	row.get(27).toString();
		TestCasePrereq = 			row.get(28).toString();
		GCCIssueCountry = 			row.get(29).toString();
		GCCPortOfArrival= 			row.get(30).toString();
		GCCArrivingFromCountry = 	row.get(31).toString();
		GCCArrivingFrom = 			row.get(32).toString();
		GCCSponsorType = 			row.get(33).toString();
		GCCSponsorNationality = 	row.get(34).toString();
		TransitComingFrom = 		row.get(35).toString();
		TransitLeavingTo = 			row.get(36).toString();
		//Comments
	}

	
	public static void fnTestDataAirline(int iTestCaseRow){ 
		//public static final String Payment_Type = "";
		//public static final String Applications_Options = "";
		
		//Test Data Sheet Columns
		Col_SrNo=					ExcelUtils.getCellData(iTestCaseRow,0);
		Col_TCJiraID=				ExcelUtils.getCellData(iTestCaseRow,1);
		Col_TestCaseName=			ExcelUtils.getCellData(iTestCaseRow,2);
		Col_ServiceName=			ExcelUtils.getCellData(iTestCaseRow,3);
		Col_EstablishmentType=		ExcelUtils.getCellData(iTestCaseRow,4);
		Col_ApplicantType=			ExcelUtils.getCellData(iTestCaseRow,5);
		Col_PersonType=				ExcelUtils.getCellData(iTestCaseRow,6);
		Col_VisitPurpose=			ExcelUtils.getCellData(iTestCaseRow,7);
		Col_Duration=				ExcelUtils.getCellData(iTestCaseRow,8);
		Col_Entry=					ExcelUtils.getCellData(iTestCaseRow,9);
		Col_PassportType=			ExcelUtils.getCellData(iTestCaseRow,10);
		Col_CurrentNationality=		ExcelUtils.getCellData(iTestCaseRow,11);
		Col_PreviousNationality=	ExcelUtils.getCellData(iTestCaseRow,12);
		Col_PassportIssueCountry=	ExcelUtils.getCellData(iTestCaseRow,13);
		Col_ResidenceNumberOfYears=	ExcelUtils.getCellData(iTestCaseRow,14);
		Col_BirthCountry=			ExcelUtils.getCellData(iTestCaseRow,15);
		Col_Gender=					ExcelUtils.getCellData(iTestCaseRow,16);
		Col_MaritalStatus=			ExcelUtils.getCellData(iTestCaseRow,17);
		Col_RelationWithSponsor=	ExcelUtils.getCellData(iTestCaseRow,18);
		Col_Religion=				ExcelUtils.getCellData(iTestCaseRow,19);
		Col_Faith=					ExcelUtils.getCellData(iTestCaseRow,20);
		Col_Education=				ExcelUtils.getCellData(iTestCaseRow,21);
		Col_Profession=				ExcelUtils.getCellData(iTestCaseRow,22);
		Col_FirstLanguage=			ExcelUtils.getCellData(iTestCaseRow,23);
		Col_Emirate=				ExcelUtils.getCellData(iTestCaseRow,24);
		Col_City=					ExcelUtils.getCellData(iTestCaseRow,25);
		Col_Area=					ExcelUtils.getCellData(iTestCaseRow,26);
		Col_CountryOutsideUAE=		ExcelUtils.getCellData(iTestCaseRow,27);
		Col_TestCasePrereq=			ExcelUtils.getCellData(iTestCaseRow,28);
		Col_GCCIssueCountry=		ExcelUtils.getCellData(iTestCaseRow,29);
		Col_GCCPortOfArrival=		ExcelUtils.getCellData(iTestCaseRow,30);
		Col_GCCArrivingFromCountry=	ExcelUtils.getCellData(iTestCaseRow,31);
		Col_GCCArrivingFrom=		ExcelUtils.getCellData(iTestCaseRow,32);
		Col_GCCSponsorType=			ExcelUtils.getCellData(iTestCaseRow,33);
		Col_GCCSponsorNationality=	ExcelUtils.getCellData(iTestCaseRow,34);
		Col_TransitComingFrom=		ExcelUtils.getCellData(iTestCaseRow,35);
		Col_TransitLeavingTo=		ExcelUtils.getCellData(iTestCaseRow,36);
		Col_Result=					ExcelUtils.getCellData(iTestCaseRow,39);
	}
	
	public static void fnTestDataTransit(int iTestCaseRow){ 
		//public static final String Payment_Type = "";
		//public static final String Applications_Options = "";
		
		//Test Data Sheet Columns
		Col_SrNo=					ExcelUtils.getCellData(iTestCaseRow,0);
		Col_TCJiraID=				ExcelUtils.getCellData(iTestCaseRow,1);
		Col_TestCaseName=			ExcelUtils.getCellData(iTestCaseRow,2);
		Col_ServiceName=			ExcelUtils.getCellData(iTestCaseRow,3);
		Col_EstablishmentType=		ExcelUtils.getCellData(iTestCaseRow,4);
		Col_ApplicantType=			ExcelUtils.getCellData(iTestCaseRow,5);
		Col_PersonType=				ExcelUtils.getCellData(iTestCaseRow,6);
		Col_VisitPurpose=			ExcelUtils.getCellData(iTestCaseRow,7);
		Col_Duration=				ExcelUtils.getCellData(iTestCaseRow,8);
		Col_Entry=					ExcelUtils.getCellData(iTestCaseRow,9);
		Col_PassportType=			ExcelUtils.getCellData(iTestCaseRow,10);
		Col_CurrentNationality=		ExcelUtils.getCellData(iTestCaseRow,11);
		Col_PreviousNationality=	ExcelUtils.getCellData(iTestCaseRow,12);
		Col_PassportIssueCountry=	ExcelUtils.getCellData(iTestCaseRow,13);
		Col_ResidenceNumberOfYears=	ExcelUtils.getCellData(iTestCaseRow,14);
		Col_BirthCountry=			ExcelUtils.getCellData(iTestCaseRow,15);
		Col_Gender=					ExcelUtils.getCellData(iTestCaseRow,16);
		Col_MaritalStatus=			ExcelUtils.getCellData(iTestCaseRow,17);
		Col_RelationWithSponsor=	ExcelUtils.getCellData(iTestCaseRow,18);
		Col_Religion=				ExcelUtils.getCellData(iTestCaseRow,19);
		Col_Faith=					ExcelUtils.getCellData(iTestCaseRow,20);
		Col_Education=				ExcelUtils.getCellData(iTestCaseRow,21);
		Col_Profession=				ExcelUtils.getCellData(iTestCaseRow,22);
		Col_FirstLanguage=			ExcelUtils.getCellData(iTestCaseRow,23);
		Col_Emirate=				ExcelUtils.getCellData(iTestCaseRow,24);
		Col_City=					ExcelUtils.getCellData(iTestCaseRow,25);
		Col_Area=					ExcelUtils.getCellData(iTestCaseRow,26);
		Col_CountryOutsideUAE=		ExcelUtils.getCellData(iTestCaseRow,27);
		Col_TestCasePrereq=			ExcelUtils.getCellData(iTestCaseRow,28);
		Col_GCCIssueCountry=		ExcelUtils.getCellData(iTestCaseRow,29);
		Col_GCCPortOfArrival=		ExcelUtils.getCellData(iTestCaseRow,30);
		Col_GCCArrivingFromCountry=	ExcelUtils.getCellData(iTestCaseRow,31);
		Col_GCCArrivingFrom=		ExcelUtils.getCellData(iTestCaseRow,32);
		Col_GCCSponsorType=			ExcelUtils.getCellData(iTestCaseRow,33);
		Col_GCCSponsorNationality=	ExcelUtils.getCellData(iTestCaseRow,34);
		Col_TransitComingFrom=		ExcelUtils.getCellData(iTestCaseRow,35);
		Col_TransitLeavingTo=		ExcelUtils.getCellData(iTestCaseRow,36);
		Col_Result=					ExcelUtils.getCellData(iTestCaseRow,39);
	}
	
	public static void fnTestDataTourism(int iTestCaseRow){ 
		Col_SrNo=					ExcelUtils.getCellData(iTestCaseRow,0);
		Col_TCJiraID=				ExcelUtils.getCellData(iTestCaseRow,1);
		Col_TestCaseName=			ExcelUtils.getCellData(iTestCaseRow,2);
		Col_ServiceName=			ExcelUtils.getCellData(iTestCaseRow,3);
		Col_EstablishmentType=		ExcelUtils.getCellData(iTestCaseRow,4);
		Col_ApplicantType=			ExcelUtils.getCellData(iTestCaseRow,5);
		Col_PersonType=				ExcelUtils.getCellData(iTestCaseRow,6);
		Col_VisitPurpose=			ExcelUtils.getCellData(iTestCaseRow,7);
		Col_Duration=				ExcelUtils.getCellData(iTestCaseRow,8);
		Col_Entry=					ExcelUtils.getCellData(iTestCaseRow,9);
		Col_PassportType=			ExcelUtils.getCellData(iTestCaseRow,10);
		Col_CurrentNationality=		ExcelUtils.getCellData(iTestCaseRow,11);
		Col_PreviousNationality=	ExcelUtils.getCellData(iTestCaseRow,12);
		Col_PassportIssueCountry=	ExcelUtils.getCellData(iTestCaseRow,13);
		Col_ResidenceNumberOfYears=	ExcelUtils.getCellData(iTestCaseRow,14);
		Col_BirthCountry=			ExcelUtils.getCellData(iTestCaseRow,15);
		Col_Gender=					ExcelUtils.getCellData(iTestCaseRow,16);
		Col_MaritalStatus=			ExcelUtils.getCellData(iTestCaseRow,17);
		Col_RelationWithSponsor=	ExcelUtils.getCellData(iTestCaseRow,18);
		Col_Religion=				ExcelUtils.getCellData(iTestCaseRow,19);
		Col_Faith=					ExcelUtils.getCellData(iTestCaseRow,20);
		Col_Education=				ExcelUtils.getCellData(iTestCaseRow,21);
		Col_Profession=				ExcelUtils.getCellData(iTestCaseRow,22);
		Col_FirstLanguage=			ExcelUtils.getCellData(iTestCaseRow,23);
		Col_Emirate=				ExcelUtils.getCellData(iTestCaseRow,24);
		Col_City=					ExcelUtils.getCellData(iTestCaseRow,25);
		Col_Area=					ExcelUtils.getCellData(iTestCaseRow,26);
		Col_CountryOutsideUAE=		ExcelUtils.getCellData(iTestCaseRow,27);
		Col_TestCasePrereq=			ExcelUtils.getCellData(iTestCaseRow,28);
		Col_GCCIssueCountry=		ExcelUtils.getCellData(iTestCaseRow,29);
		Col_GCCPortOfArrival=		ExcelUtils.getCellData(iTestCaseRow,30);
		Col_GCCArrivingFromCountry=	ExcelUtils.getCellData(iTestCaseRow,31);
		Col_GCCArrivingFrom=		ExcelUtils.getCellData(iTestCaseRow,32);
		Col_GCCSponsorType=			ExcelUtils.getCellData(iTestCaseRow,33);
		Col_GCCSponsorNationality=	ExcelUtils.getCellData(iTestCaseRow,34);
		Col_TransitComingFrom=		ExcelUtils.getCellData(iTestCaseRow,35);
		Col_TransitLeavingTo=		ExcelUtils.getCellData(iTestCaseRow,36);
		Col_Result=					ExcelUtils.getCellData(iTestCaseRow,39);
	}
	
	public static void fnTestDataVisit(int iTestCaseRow){ 
		Col_SrNo=					ExcelUtils.getCellData(iTestCaseRow,0);
		Col_TCJiraID=				ExcelUtils.getCellData(iTestCaseRow,1);
		Col_TestCaseName=			ExcelUtils.getCellData(iTestCaseRow,2);
		Col_ServiceName=			ExcelUtils.getCellData(iTestCaseRow,3);
		Col_EstablishmentType=		ExcelUtils.getCellData(iTestCaseRow,4);
		Col_ApplicantType=			ExcelUtils.getCellData(iTestCaseRow,5);
		Col_PersonType=				ExcelUtils.getCellData(iTestCaseRow,6);
		Col_VisitPurpose=			ExcelUtils.getCellData(iTestCaseRow,7);
		Col_Duration=				ExcelUtils.getCellData(iTestCaseRow,8);
		Col_Entry=					ExcelUtils.getCellData(iTestCaseRow,9);
		Col_PassportType=			ExcelUtils.getCellData(iTestCaseRow,10);
		Col_CurrentNationality=		ExcelUtils.getCellData(iTestCaseRow,11);
		Col_PreviousNationality=	ExcelUtils.getCellData(iTestCaseRow,12);
		Col_PassportIssueCountry=	ExcelUtils.getCellData(iTestCaseRow,13);
		Col_ResidenceNumberOfYears=	ExcelUtils.getCellData(iTestCaseRow,14);
		Col_BirthCountry=			ExcelUtils.getCellData(iTestCaseRow,15);
		Col_Gender=					ExcelUtils.getCellData(iTestCaseRow,16);
		Col_MaritalStatus=			ExcelUtils.getCellData(iTestCaseRow,17);
		Col_RelationWithSponsor=	ExcelUtils.getCellData(iTestCaseRow,18);
		Col_Religion=				ExcelUtils.getCellData(iTestCaseRow,19);
		Col_Faith=					ExcelUtils.getCellData(iTestCaseRow,20);
		Col_Education=				ExcelUtils.getCellData(iTestCaseRow,21);
		Col_Profession=				ExcelUtils.getCellData(iTestCaseRow,22);
		Col_FirstLanguage=			ExcelUtils.getCellData(iTestCaseRow,23);
		Col_Emirate=				ExcelUtils.getCellData(iTestCaseRow,24);
		Col_City=					ExcelUtils.getCellData(iTestCaseRow,25);
		Col_Area=					ExcelUtils.getCellData(iTestCaseRow,26);
		Col_CountryOutsideUAE=		ExcelUtils.getCellData(iTestCaseRow,27);
		Col_TestCasePrereq=			ExcelUtils.getCellData(iTestCaseRow,28);
		Col_GCCIssueCountry=		ExcelUtils.getCellData(iTestCaseRow,29);
		Col_GCCPortOfArrival=		ExcelUtils.getCellData(iTestCaseRow,30);
		Col_GCCArrivingFromCountry=	ExcelUtils.getCellData(iTestCaseRow,31);
		Col_GCCArrivingFrom=		ExcelUtils.getCellData(iTestCaseRow,32);
		Col_GCCSponsorType=			ExcelUtils.getCellData(iTestCaseRow,33);
		Col_GCCSponsorNationality=	ExcelUtils.getCellData(iTestCaseRow,34);
		Col_TransitComingFrom=		ExcelUtils.getCellData(iTestCaseRow,35);
		Col_TransitLeavingTo=		ExcelUtils.getCellData(iTestCaseRow,36);
		Col_Result=					ExcelUtils.getCellData(iTestCaseRow,39);
	}
	
	public static void fnTestDataGCC(int iTestCaseRow){ 
		Col_SrNo=					ExcelUtils.getCellData(iTestCaseRow,0);
		Col_TCJiraID=				ExcelUtils.getCellData(iTestCaseRow,1);
		Col_TestCaseName=			ExcelUtils.getCellData(iTestCaseRow,2);
		Col_ServiceName=			ExcelUtils.getCellData(iTestCaseRow,3);
		Col_EstablishmentType=		ExcelUtils.getCellData(iTestCaseRow,4);
		Col_ApplicantType=			ExcelUtils.getCellData(iTestCaseRow,5);
		Col_PersonType=				ExcelUtils.getCellData(iTestCaseRow,6);
		Col_VisitPurpose=			ExcelUtils.getCellData(iTestCaseRow,7);
		Col_Duration=				ExcelUtils.getCellData(iTestCaseRow,8);
		Col_Entry=					ExcelUtils.getCellData(iTestCaseRow,9);
		Col_PassportType=			ExcelUtils.getCellData(iTestCaseRow,10);
		Col_CurrentNationality=		ExcelUtils.getCellData(iTestCaseRow,11);
		Col_PreviousNationality=	ExcelUtils.getCellData(iTestCaseRow,12);
		Col_PassportIssueCountry=	ExcelUtils.getCellData(iTestCaseRow,13);
		Col_ResidenceNumberOfYears=	ExcelUtils.getCellData(iTestCaseRow,14);
		Col_BirthCountry=			ExcelUtils.getCellData(iTestCaseRow,15);
		Col_Gender=					ExcelUtils.getCellData(iTestCaseRow,16);
		Col_MaritalStatus=			ExcelUtils.getCellData(iTestCaseRow,17);
		Col_RelationWithSponsor=	ExcelUtils.getCellData(iTestCaseRow,18);
		Col_Religion=				ExcelUtils.getCellData(iTestCaseRow,19);
		Col_Faith=					ExcelUtils.getCellData(iTestCaseRow,20);
		Col_Education=				ExcelUtils.getCellData(iTestCaseRow,21);
		Col_Profession=				ExcelUtils.getCellData(iTestCaseRow,22);
		Col_FirstLanguage=			ExcelUtils.getCellData(iTestCaseRow,23);
		Col_Emirate=				ExcelUtils.getCellData(iTestCaseRow,24);
		Col_City=					ExcelUtils.getCellData(iTestCaseRow,25);
		Col_Area=					ExcelUtils.getCellData(iTestCaseRow,26);
		Col_CountryOutsideUAE=		ExcelUtils.getCellData(iTestCaseRow,27);
		Col_TestCasePrereq=			ExcelUtils.getCellData(iTestCaseRow,28);
		Col_GCCIssueCountry=		ExcelUtils.getCellData(iTestCaseRow,29);
		Col_GCCPortOfArrival=		ExcelUtils.getCellData(iTestCaseRow,30);
		Col_GCCArrivingFromCountry=	ExcelUtils.getCellData(iTestCaseRow,31);
		Col_GCCArrivingFrom=		ExcelUtils.getCellData(iTestCaseRow,32);
		Col_GCCSponsorType=			ExcelUtils.getCellData(iTestCaseRow,33);
		Col_GCCSponsorNationality=	ExcelUtils.getCellData(iTestCaseRow,34);
		Col_TransitComingFrom=		ExcelUtils.getCellData(iTestCaseRow,35);
		Col_TransitLeavingTo=		ExcelUtils.getCellData(iTestCaseRow,36);
		Col_Result=					ExcelUtils.getCellData(iTestCaseRow,39);
	}
	
	public static void fnTestDataWork(int iTestCaseRow){ 
		Col_SrNo=					ExcelUtils.getCellData(iTestCaseRow,0);
		Col_TCJiraID=				ExcelUtils.getCellData(iTestCaseRow,1);
		Col_TestCaseName=			ExcelUtils.getCellData(iTestCaseRow,2);
		Col_ServiceName=			ExcelUtils.getCellData(iTestCaseRow,3);
		Col_EstablishmentType=		ExcelUtils.getCellData(iTestCaseRow,4);
		Col_ApplicantType=			ExcelUtils.getCellData(iTestCaseRow,5);
		Col_PersonType=				ExcelUtils.getCellData(iTestCaseRow,6);
		Col_VisitPurpose=			ExcelUtils.getCellData(iTestCaseRow,7);
		Col_Duration=				ExcelUtils.getCellData(iTestCaseRow,8);
		Col_Entry=					ExcelUtils.getCellData(iTestCaseRow,9);
		Col_PassportType=			ExcelUtils.getCellData(iTestCaseRow,10);
		Col_CurrentNationality=		ExcelUtils.getCellData(iTestCaseRow,11);
		Col_PreviousNationality=	ExcelUtils.getCellData(iTestCaseRow,12);
		Col_PassportIssueCountry=	ExcelUtils.getCellData(iTestCaseRow,13);
		Col_ResidenceNumberOfYears=	ExcelUtils.getCellData(iTestCaseRow,14);
		Col_BirthCountry=			ExcelUtils.getCellData(iTestCaseRow,15);
		Col_Gender=					ExcelUtils.getCellData(iTestCaseRow,16);
		Col_MaritalStatus=			ExcelUtils.getCellData(iTestCaseRow,17);
		Col_RelationWithSponsor=	ExcelUtils.getCellData(iTestCaseRow,18);
		Col_Religion=				ExcelUtils.getCellData(iTestCaseRow,19);
		Col_Faith=					ExcelUtils.getCellData(iTestCaseRow,20);
		Col_Education=				ExcelUtils.getCellData(iTestCaseRow,21);
		Col_Profession=				ExcelUtils.getCellData(iTestCaseRow,22);
		Col_FirstLanguage=			ExcelUtils.getCellData(iTestCaseRow,23);
		Col_Emirate=				ExcelUtils.getCellData(iTestCaseRow,24);
		Col_City=					ExcelUtils.getCellData(iTestCaseRow,25);
		Col_Area=					ExcelUtils.getCellData(iTestCaseRow,26);
		Col_CountryOutsideUAE=		ExcelUtils.getCellData(iTestCaseRow,27);
		Col_TestCasePrereq=			ExcelUtils.getCellData(iTestCaseRow,28);
		Col_GCCIssueCountry=		ExcelUtils.getCellData(iTestCaseRow,29);
		Col_GCCPortOfArrival=		ExcelUtils.getCellData(iTestCaseRow,30);
		Col_GCCArrivingFromCountry=	ExcelUtils.getCellData(iTestCaseRow,31);
		Col_GCCArrivingFrom=		ExcelUtils.getCellData(iTestCaseRow,32);
		Col_GCCSponsorType=			ExcelUtils.getCellData(iTestCaseRow,33);
		Col_GCCSponsorNationality=	ExcelUtils.getCellData(iTestCaseRow,34);
		Col_TransitComingFrom=		ExcelUtils.getCellData(iTestCaseRow,35);
		Col_TransitLeavingTo=		ExcelUtils.getCellData(iTestCaseRow,36);
		Col_Result=					ExcelUtils.getCellData(iTestCaseRow,39);
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
