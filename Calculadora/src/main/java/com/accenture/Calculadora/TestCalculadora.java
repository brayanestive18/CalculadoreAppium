package com.accenture.Calculadora;

import org.testng.annotations.Test;
import org.testng.AssertJUnit;
//import org.testng.annotations.Test;
//import org.testng.AssertJUnit;
//import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;
import com.aventstack.extentreports.reporter.configuration.ChartLocation;
import com.aventstack.extentreports.reporter.configuration.Theme;

//import java.awt.AWTException;
//import java.awt.Rectangle;
//import java.awt.Robot;
//import java.awt.Toolkit;
//import java.awt.image.BufferedImage;
import java.io.File;
//import org.testng.annotations.Test;
//import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
//import java.util.Arrays;
//import java.io.InputStream;
import java.util.Iterator;
import java.util.Date;

import java.net.MalformedURLException;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.remote.DesiredCapabilities;
//import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
//import org.testng.annotations.Test;

import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;
import io.appium.java_client.android.AndroidDriver;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;	
import com.aventstack.extentreports.markuputils.MarkupHelper;
//import com.aventstack.extentreports.reporter.ExtentHtmlReporter;
//import com.aventstack.extentreports.reporter.configuration.ChartLocation;
//import com.aventstack.extentreports.reporter.configuration.Theme;


public class TestCalculadora {

	private ArrayList<String> operacion = new ArrayList<>();
	
	private ExtentHtmlReporter htmlReporter;
	private ExtentReports extent;
    private static ExtentTest test, testSum, testRes, testMul, testDiv, testResul;
	
	public static AppiumDriver<MobileElement> driver; //Este driver es el que contralara los eventos de la automatizacion
	DesiredCapabilities capabilities = new DesiredCapabilities(); //caracteristicas de la automatizacion
	//private HSSFWorkbook worbook;

	@BeforeMethod
	public void setUpAppium() throws MalformedURLException, InterruptedException {
		
		String nombreArchivo  = "TestCal.xlsx";
		String rutaArchivo = "C:\\Users\\brayan.diaz\\Documents\\" + nombreArchivo ;
		//String hoja = "Test";
		
		
		System.out.println("-------------------------- Start Test --------------------------\n" + rutaArchivo);

		//Configurar el HTML de extent Report
		
		htmlReporter = new ExtentHtmlReporter("C:\\Users\\brayan.diaz\\Documents\\ReportCalculator.html");
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
		
        test = extent.createTest("Demo", "This test");
        AssertJUnit.assertTrue(true);
        
        
        testSum = extent.createTest("Suma", "This test");
  		AssertJUnit.assertTrue(true);
        testRes = extent.createTest("Resta", "This test");
  		AssertJUnit.assertTrue(true);
        testMul = extent.createTest("Multiplicaciones", "This test");
  		AssertJUnit.assertTrue(true);
        testDiv = extent.createTest("Division", "This test");
  		AssertJUnit.assertTrue(true);
        testResul = extent.createTest("Resultado", "This test");
  		AssertJUnit.assertTrue(true);
        
		FileInputStream excelStream = null;
		
		try {
			
			excelStream = new FileInputStream(rutaArchivo);
			
			XSSFWorkbook worbook = new XSSFWorkbook(excelStream);
			//obtener la hoja que se va leer
			XSSFSheet sheet1 = worbook.getSheetAt(0);
			//obtener todas las filas de la hoja excel
			Iterator<Row> rowIterator = sheet1.iterator();
 
			//HSSFRow hssfRow;
			//HSSFCell cell1;
			
			// Obtengo el número de filas ocupadas en la hoja
            //int rows = sheet1.getLastRowNum();
            //int cols = 0;
            
            //String cellValue;
			
			Row row;
			// se recorre cada fila hasta el final

			System.out.println("-------------------------- Archivo de Excel --------------------------");
			String save;
			while (rowIterator.hasNext()) {
				row = rowIterator.next();
				
				Iterator<Cell> cellIterator = row.cellIterator();
				
				
				Cell celda;
				//System.out.println("- ");
				while (cellIterator.hasNext()) {
					celda = cellIterator.next();
					switch(celda.getCellType()) {
					//switch(celda.getCellTypeEnum().toString()) {
					//case "NUMERIC"
					case Cell.CELL_TYPE_NUMERIC:
						if( DateUtil.isCellDateFormatted(celda) ){
							//System.out.println(celda.getDateCellValue());
						} else{
							//System.out.println(celda.getNumericCellValue());
							double num = celda.getNumericCellValue();
							
							int num1 = (int) num;
							double num2 = (double) num1;
							// Para quitar el punto y el cero en los numeros enteros
							if ((num - num2) == 0) {
								/*System.out.println("----------------");
								System.out.println(num1);
								System.out.println(num2);
								System.out.println(num - num2);
								num = (double)num1;
								System.out.println(num);*/
								save = Integer.toString(num1);
							} else {
								save = Double.toString(num);
							}
							String[] sep = save.split("");
							for (int i = 0; i < sep.length; i++) {
								//System.out.println(sep[i]);
																
								operacion.add(sep[i]);
								
							}
							
							//operacion.add(num);
						}
						break;
					//case "STRING":
					case Cell.CELL_TYPE_STRING:
						//System.out.println(celda.getStringCellValue());
						operacion.add(celda.getStringCellValue());
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						//System.out.println(celda.getBooleanCellValue());
						break;
					}
				}
				
			}
			
			worbook.close();
			
		} catch (Exception e) {
			e.getMessage();
			System.out.println("-------------------------- No encontro - Excel --------------------------\n" + e);
		}
		
		for(String ope:operacion) {
			System.out.println(ope);
			//System.out.println(" ** ");
		}
		
		//Configuración para conectar con el Appium
		String packagename = "de.underflow.calc"; //Paquete principal de la aplicacion a automatizar
		//String URL = "http://172.16.1.55:4760/wd/hub"; //IP y puerto de Appium
		String URL = "http://127.0.0.1:4723/wd/hub";
		String activityname = "de.underflow.calc.CalculatorMainActivity"; //Nombre de la actividad (o vista) en donde empezara la automatizacion
		capabilities.setCapability("deviceName", "Samsung S6"); //No es obligatorio que este nombre coincida
		//capabilities.setCapability("udid", "172.16.1.65:8526"); //Serial del dispositivo, se obtiene activando la depuración USB y con el comando adb devices
		capabilities.setCapability("udid", "03157df36237c113");
		capabilities.setCapability("platformVersion", "7.0"); //No es obligatorio que la version coincida
		capabilities.setCapability("platformName", "Android"); //Nombre del sistema operativo
		capabilities.setCapability("appPackage", packagename);
		capabilities.setCapability("appActivity", activityname);
		driver = new AndroidDriver<MobileElement>(new URL(URL), capabilities);
		driver.manage().timeouts().implicitlyWait(80, TimeUnit.SECONDS);
	}

	@AfterTest
	public void CleanUpAppium() {
		driver.quit();
	}

	@Test
	public void mytest() throws InterruptedException {
		
		ArrayList<String> resultados = new ArrayList<>();
		String path;
		
		try {
			Thread.sleep(1000);
			String opera = "|";
			
			for(String lec:operacion) {
				
				/*if (lec == "=") {
					
					pos = operacion.indexOf("=");
					System.out.println("Pos: " + pos);
					operacion.set(pos, "|");
					
					System.out.println(lec);
					capturarPantalla();
				}*/
				
				switch (lec) {
				case "1":
					MobileElement num1 = driver.findElement(By.id("de.underflow.calc:id/One"));
					num1.click();
					break;
					
				case "2":
					MobileElement num2 = driver.findElement(By.id("de.underflow.calc:id/Two"));
					num2.click();
					break;
					
				case "3":
					MobileElement num3 = driver.findElement(By.id("de.underflow.calc:id/Three"));
					num3.click();
					break;
					
				case "4":
					MobileElement num4 = driver.findElement(By.id("de.underflow.calc:id/Four"));
					num4.click();
					break;
					
				case "5":
					MobileElement num5 = driver.findElement(By.id("de.underflow.calc:id/Five"));
					num5.click();
					break;
					
				case "6":
					MobileElement num6 = driver.findElement(By.id("de.underflow.calc:id/Six"));
					num6.click();
					break;
					
				case "7":
					MobileElement num7 = driver.findElement(By.id("de.underflow.calc:id/Seven"));
					num7.click();
					break;
					
				case "8":
					MobileElement num8 = driver.findElement(By.id("de.underflow.calc:id/Eight"));
					num8.click();
					break;
					
				case "9":
					MobileElement num9 = driver.findElement(By.id("de.underflow.calc:id/Nine"));
					num9.click();
					break;

				case "0":
					MobileElement num0 = driver.findElement(By.id("de.underflow.calc:id/Zero"));
					num0.click();
					break;
				
				case "+":
					MobileElement mas = driver.findElement(By.id("de.underflow.calc:id/Plus"));
					//test.log(Status.PASS, MarkupHelper.createLabel("Operación Suma",ExtentColor.GREEN));
					mas.click();
					opera = "+";
					break;
				
				case "-":
					MobileElement minus = driver.findElement(By.id("de.underflow.calc:id/Minus"));
					minus.click();
					opera = "-";
					break;
				
				case "*":
					MobileElement mul = driver.findElement(By.id("de.underflow.calc:id/Multiply"));
					mul.click();
					opera = "*";
					break;
				
				case "/":
					MobileElement div = driver.findElement(By.id("de.underflow.calc:id/Divide"));
					div.click();
					opera = "/";
					break;
				
				case "=":
					path = capturarPantalla();
					MobileElement igual = driver.findElement(By.id("de.underflow.calc:id/Equals"));
					igual.click();
					MobileElement result = driver.findElementById("de.underflow.calc:id/Result");
					//
					String textResult = result.getText();
					//EditExcel(textResult);
					resultados.add(textResult);
					path = "C:\\Users\\brayan.diaz\\eclipse-workspace\\Calculadora\\Screenshots\\" + path;
					System.out.println("Path: " + path);
					System.out.println("Result: " + textResult);
					//Hacer el log
					loadLog(opera, path);
					path = capturarPantalla();
					opera = "=";
					path = "C:\\Users\\brayan.diaz\\eclipse-workspace\\Calculadora\\Screenshots\\" + path;
					loadLog(opera, path);
					break;
					
				case ".":
					MobileElement dot = driver.findElement(By.id("de.underflow.calc:id/Dot"));
					dot.click();
					break;
					
				default:
					break;
				}
			}
			
			EditExcel(resultados);
			extent.flush();
			
			/*
			// hacer click boton2
			MobileElement num2 = driver.findElement(By.id("de.underflow.calc:id/Two"));
			num2.click();

			// Hacer click boton7
			MobileElement num7 = driver.findElement(By.id("de.underflow.calc:id/Seven"));
			num7.click();
			
			// hacer click boton mas( +)
			MobileElement mas = driver.findElement(By.id("de.underflow.calc:id/Plus"));
			mas.click();

			// hacer click boton9
			MobileElement num9 = driver.findElement(By.id("de.underflow.calc:id/Nine"));
			num9.click();

			// hacer click boton mas( +)
			MobileElement mas1 = driver.findElement(By.id("de.underflow.calc:id/Plus"));
			mas1.click();

			// hacer click boton8
			MobileElement num8 = driver.findElement(By.id("de.underflow.calc:id/Eight"));
			num8.click();

			// hacer click boton igual
			MobileElement igual = driver.findElement(By.id("de.underflow.calc:id/Equals"));
			igual.click();

			// hacer click boton dividir
			MobileElement dividido = driver.findElement(By.id("de.underflow.calc:id/Divide"));
			dividido.click();

			// hacer click boton 2
			MobileElement numdos = driver.findElement(By.id("de.underflow.calc:id/Two"));
			numdos.click();

			// hacer click boton =
			MobileElement igual2 = driver.findElement(By.id("de.underflow.calc:id/Equals"));
			igual2.click();

			MobileElement result = driver.findElementById("de.underflow.calc:id/Result");
			String textResult = result.getText();
			System.out.println("Result: " + textResult);
			*/
		} catch (Exception e) {
			System.out.println("Se presento Excepción " + e);
		}
	}
	
	public void EditExcel(ArrayList<String> value) {
		
		String nombreArchivo  = "TestCal.xlsx";
		String rutaArchivo = "C:\\Users\\brayan.diaz\\Documents\\" + nombreArchivo ;
		
		FileInputStream file = null;
		
		try {
			
			file = new FileInputStream(rutaArchivo);
			
			XSSFWorkbook book = new XSSFWorkbook(file);
			//obtener la hoja que se va leer
			XSSFSheet sheet_1 = book.getSheetAt(0);
			
			for(int i = 0; i < value.size(); i++) {
				XSSFRow fila = sheet_1.getRow(i);
				if (fila == null) {
					fila = sheet_1.createRow(i);
				}
				
				XSSFCell celda = fila.getCell(5); 
				
				if(celda == null) {
					celda = fila.createCell(5);
				}
				
				//celda.setCellValue(Double.parseDouble(value.get(i)));
				celda.setCellValue(value.get(i));
				
				
			}
			
			file.close();
			
			FileOutputStream output = new FileOutputStream(rutaArchivo);
			book.write(output);
			output.close();
			
			
		} catch (Exception e) {
			test.log(Status.ERROR, MarkupHelper.createLabel("NO ENCONTRO EL EXCEL",ExtentColor.RED));
			e.getMessage();
			System.out.println("-------------------------- No encontro - Excel --------------------------\n" + e);
		}
	}
	
	  public static String capturarPantalla() {
		  String dirScreen = "Screenshots";
		  File screenFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		  DateFormat dateFormat = new SimpleDateFormat("dd-mm-yy_hh_mm_ssaa");
		  new File(dirScreen).mkdirs();
		  String destFile = dateFormat.format(new Date()) + ".png";
		  File ruta = new File(dirScreen + "/" + destFile);
		  try {
			  FileUtils.copyFile(screenFile, ruta);
		  } catch (Exception e) {
			  test.log(Status.ERROR, MarkupHelper.createLabel("NO SCREENSHOT",ExtentColor.RED));
			  e.printStackTrace();
		  }
		  //return ruta.toString();
		  return destFile;
	  }
	  
	  public void loadLog(String operator, String path) throws IOException {
		  switch (operator) {
		  	case "+":
		  		//test = extent.createTest("Suma", "This test");
		  		//AssertJUnit.assertTrue(true);
		  		testSum.log(Status.PASS, MarkupHelper.createLabel("Operacion Suma",ExtentColor.GREEN));
		  		testSum.addScreenCaptureFromPath(path);
			break;

		  	case "-":
		  		//test = extent.createTest("Resta", "This test");
		  		//AssertJUnit.assertTrue(true);
		  		testRes.log(Status.PASS, MarkupHelper.createLabel("Operacion Resta",ExtentColor.GREEN));
		  		testRes.addScreenCaptureFromPath(path);
			break;
			
		  	case "*":
		  		//test = extent.createTest("Multiplicacion", "This test");
		  		//AssertJUnit.assertTrue(true);
		  		testMul.log(Status.PASS, MarkupHelper.createLabel("Operacion Multiplicacion",ExtentColor.GREEN));
		  		testMul.addScreenCaptureFromPath(path);
			break;
			
		  	case "/":
		  		//test = extent.createTest("Division", "This test");
		  		//AssertJUnit.assertTrue(true);
		  		testDiv.log(Status.PASS, MarkupHelper.createLabel("Operacion Division",ExtentColor.GREEN));
		  		testDiv.addScreenCaptureFromPath(path);
			break;
			
		  	case "=":
		  		//test = extent.createTest("Resultados", "This test");
		  		//AssertJUnit.assertTrue(true);
		  		//testResul.log(Status.INFO, "Operacion: " + operator); 
		  		testResul.log(Status.PASS, MarkupHelper.createLabel("Resultados de la operaciones",ExtentColor.GREEN));
		  		testResul.addScreenCaptureFromPath(path);
			break;
			
		  	default:
			break;
		}
	  }
	
}

