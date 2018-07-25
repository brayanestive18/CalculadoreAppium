package com.accenture.Calculadora;

import org.testng.Assert;
import org.testng.ITestResult;
import org.testng.SkipException;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;	
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;
import com.aventstack.extentreports.reporter.configuration.ChartLocation;
import com.aventstack.extentreports.reporter.configuration.Theme;

public class GenerateExtentReport {
	
	ExtentHtmlReporter htmlReporter;
	ExtentReports extent;
    ExtentTest test;
    
    @BeforeTest
    public void startReport() {
    	htmlReporter = new ExtentHtmlReporter("C:\\Users\\brayan.diaz\\Documents\\MyOwnReport.html");
    	extent = new ExtentReports();
        extent.attachReporter(htmlReporter);
        
        extent.setSystemInfo("OS", "Android");
        extent.setSystemInfo("Host Name", "brayan.diaz");
        extent.setSystemInfo("User Name", "Brayan Diaz");
         
        htmlReporter.config().setChartVisibilityOnOpen(true);
        htmlReporter.config().setDocumentTitle("AutomationTesting.in Demo Report");
        htmlReporter.config().setReportName("Reporte Brayan");
        htmlReporter.config().setTestViewChartLocation(ChartLocation.TOP);
        htmlReporter.config().setTheme(Theme.DARK);
        
    }
    
    @Test
    public void demoTestPass()
    {
        test = extent.createTest("demoTestPass", "This test will demonstrate the PASS test case");
        Assert.assertTrue(true);
    }
     
    @Test
    public void demoTestFail()
    {
        test = extent.createTest("demoTestFail", "This test will demonstrate the FAIL test case");
        Assert.assertTrue(true);
    }
     
    @Test
    public void demoTestSkip()
    {
        test = extent.createTest("demoTestSkip", "This test will demonstrate the SKIP test case");
        throw new SkipException("This test case not ready for execution");
    }
     
    @AfterMethod
    public void getResult(ITestResult result)
    {
        if(result.getStatus() == ITestResult.FAILURE)
        {
            test.log(Status.FAIL, MarkupHelper.createLabel(result.getName()+" Test case FAILED due to below issues:", ExtentColor.RED));
            test.fail(result.getThrowable());
        }
        else if(result.getStatus() == ITestResult.SUCCESS)
        {
            test.log(Status.PASS, MarkupHelper.createLabel(result.getName()+" Test Case PASSED", ExtentColor.GREEN));
            test.log(Status.PASS, MarkupHelper.createLabel("Hola Mundo",ExtentColor.GREEN));	
        }
        else
        {
            test.log(Status.SKIP, MarkupHelper.createLabel(result.getName()+" Test Case SKIPPED", ExtentColor.ORANGE));
            test.skip(result.getThrowable());
        }
    }
     
    @AfterTest
    public void tearDown()
    {
        extent.flush();
    }
    
}
