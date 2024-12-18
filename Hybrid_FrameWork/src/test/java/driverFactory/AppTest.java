 package driverFactory;

import java.util.Iterator;

import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;

import commonFunctions.FunctionLibrary;
import utilities.ExcelFileUtil;

public class AppTest {
String inputpath ="./FileInput/DataEngine_lyst1733124900878.xlsx";
String outputpath = "./FileOutput/HybridResults.xlsx";
String TCsheet = "MasterTestCases";
ExtentReports reports;
ExtentTest logger;
WebDriver driver;
@Test
public void StartTest() throws Throwable
{
	String Module_status = "";
	String Module_New = "";
	//create object for excelfile util class
	ExcelFileUtil xl = new ExcelFileUtil(inputpath);
	//count no of rows in TCsheet
	int rc = xl.rowCount(TCsheet);
	//iterate all rows
	for(int i=1;i<=rc;i++)
	{
		if(xl.getcellData(TCsheet, i, 2).equalsIgnoreCase("Y"))
		{
			//read  corresponding sheet from TCsheet
			String TCModule = xl.getcellData(TCsheet, i, 1);
			//iterate all rows in TCModule
			for(int j=1;j<=xl.rowCount(TCModule);j++)
			{
				//read each cell from TCmodule
				String Description = xl.getcellData(TCModule, j, 0);
				String ObjectType = xl.getcellData(TCModule, j, 1);
				String LocatorType = xl.getcellData(TCModule, j, 2);
				String LocatorValue = xl.getcellData(TCModule, j, 3);
				String TestData = xl.getcellData(TCModule, j, 4);
				try {
			if(ObjectType.equalsIgnoreCase("startBrowser"))
					{
						driver = FunctionLibrary.startBrowser();
					}
			if(ObjectType.equalsIgnoreCase("openUrl"))
					{
						FunctionLibrary.openUrl();
					}
			if(ObjectType.equalsIgnoreCase("waitforElement"))
					{
						FunctionLibrary.waitForElement(LocatorType, LocatorValue, TestData);
					}
			if(ObjectType.equalsIgnoreCase("typeAction"))
					{
						FunctionLibrary.typeAction(LocatorType, LocatorValue, TestData);
					}
			if(ObjectType.equalsIgnoreCase("clickAction"))
					{
						FunctionLibrary.clickAction(LocatorType, LocatorValue);
					}
			if(ObjectType.equalsIgnoreCase("validateTitle"))
					{
						FunctionLibrary.validateTitle(TestData);
					}
			if(ObjectType.equalsIgnoreCase("closeBrowser"))
					{
						FunctionLibrary.closeBrowser();
					}
			//write as pass into status cell in TCModule 5th column
			xl.setcelldata(TCModule, j, 5, "Pass", outputpath);
			Module_status = "True";
				}catch 
				(Exception e) {
					System.out.println(e.getMessage());
					//write as Fail into status cell in TCModule 5th column
					xl.setcelldata(TCModule, j, 5, "Fail", outputpath);
					Module_New = "False";
					}
				if(Module_status.equalsIgnoreCase("True"))
				{
					//write as pass into TCSheet
					xl.setcelldata(TCsheet, i, 3, "Pass", outputpath);
				}
				if(Module_New.equalsIgnoreCase("False"))
				{
					//write as Fail into TCSheet
					xl.setcelldata(TCsheet, i, 3, "Fail", outputpath);
				}
			}
		}
		else
		{
			//write as blocked into TCsheet which testcase flag to N
		xl.setcelldata(TCsheet, i, 3, "Blocked", outputpath);
		}
	}
}
}
