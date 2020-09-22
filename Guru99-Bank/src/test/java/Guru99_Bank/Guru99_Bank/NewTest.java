package Guru99_Bank.Guru99_Bank;

import org.testng.annotations.Test;
import org.testng.asserts.SoftAssert;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.BeforeClass;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.AfterSuite;

public class NewTest {
	
	WebDriver driver;
	

	@BeforeTest
	public void setup() {
		System.setProperty("webdriver.gecko.driver",Util.FIREFOX_PATH);
		driver = new FirefoxDriver();
		driver.navigate().to(Util.BASE_URL);
	}
		
	
	public void readExcel(String filePath, String fileName, String sheetName) throws IOException, InterruptedException {
		File file = new File(filePath+"\\"+fileName);
		FileInputStream fs = new FileInputStream(file);
		Workbook workBook = null;
		String fileExtensionName = fileName.substring(fileName.indexOf("."));
		if (fileExtensionName.equals(".xlsx")) {
			workBook = new XSSFWorkbook(fs); 
		}
		if (fileExtensionName.equals(".xls")) {
			workBook = new HSSFWorkbook(fs);
		}
		Sheet sheet = workBook.getSheet(sheetName);
		for (int i = 1; i <= sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
				Cell c = row.getCell(0);
				String value = c.getStringCellValue();
				driver.findElement(By.xpath("//input[@name='uid']")).sendKeys(value);
				driver.findElement(By.xpath("//input[@name='password']")).sendKeys(row.getCell(1).getStringCellValue());
				driver.findElement(By.xpath("//input[@type='submit']")).click();
				Thread.sleep(5000);
				String actualTitle = driver.getTitle();
				String expectedTitle = "Home Page";
				SoftAssert sassert = new SoftAssert();
				sassert.assertEquals(actualTitle, expectedTitle);
				sassert.assertAll();
				break; 	
		}
		
	}
	
	@Test
	public static void callingFunction() throws IOException, InterruptedException {
		NewTest newTest = new NewTest();
		String filePath = "C:\\Users\\Sitesh\\Desktop";
		newTest.readExcel(filePath, "Book2.xlsx", "Sheet1");
	}

	
	@AfterTest
	public void close() {
		driver.close();
	}
}
