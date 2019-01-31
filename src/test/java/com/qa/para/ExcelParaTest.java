package com.qa.para;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriver;





@RunWith(Parameterized.class)
public class ExcelParaTest {

	@Parameters
	public static Collection<Object[]> data() throws IOException {
		FileInputStream file = new FileInputStream("C:\\Users\\Admin\\Documents\\DemoSiteDDT.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		Object[][] ob = new Object[sheet.getPhysicalNumberOfRows()-1][4];
		
//		Reading
		for (int rowNum = 1; rowNum < sheet.getPhysicalNumberOfRows(); rowNum++) {
				ob[rowNum-1][0] = sheet.getRow(rowNum).getCell(0).getStringCellValue();
				ob[rowNum-1][1] = sheet.getRow(rowNum).getCell(1).getStringCellValue();
				ob[rowNum-1][2] = sheet.getRow(rowNum).getCell(2).getStringCellValue();
				ob[rowNum-1][3] = rowNum;
			}
		return Arrays.asList(ob);
		}
	
	private String username;
	private String password;
	private String expected;
	private int rowNum;
	private WebDriver driver;
	
	public ExcelParaTest(String username, String password, String expected, int rowNum) {
		this.username = username;
		this.password = password;
		this.expected = expected;
		this.rowNum = rowNum;
	}
	
	@Before
	public void setup() {
		System.setProperty("phantomjs.binary.path", Constants.PHANTOMDRIVER);
		driver = new PhantomJSDriver();
	}
	@Test
	public void login() throws Exception {
		// testing logic
		
		// Reading
		
		
			driver.get("http://thedemosite.co.uk/addauser.php");
			WebElement textbox = driver.findElement(By.xpath("/html/body/table/tbody/tr/td[1]/form/div/center/table/tbody/tr/td[1]/div/center/table/tbody/tr[1]/td[2]/p/input"));
			textbox.sendKeys(username);
			WebElement textbox2 = driver.findElement(By.xpath("/html/body/table/tbody/tr/td[1]/form/div/center/table/tbody/tr/td[1]/div/center/table/tbody/tr[2]/td[2]/p/input"));
			textbox2.sendKeys(password);
			WebElement button = driver.findElement(By.xpath("/html/body/table/tbody/tr/td[1]/form/div/center/table/tbody/tr/td[1]/div/center/table/tbody/tr[3]/td[2]/p/input"));
			button.click();
			
			
			driver.get("http://thedemosite.co.uk/login.php");
			WebElement usernamel = driver.findElement(By.xpath("/html/body/table/tbody/tr/td[1]/form/div/center/table/tbody/tr/td[1]/table/tbody/tr[1]/td[2]/p/input"));
			usernamel.sendKeys(username);
			WebElement passwordl = driver.findElement(By.xpath("/html/body/table/tbody/tr/td[1]/form/div/center/table/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/p/input"));
			passwordl.sendKeys(password);
			
			WebElement button2 = driver.findElement(By.xpath("/html/body/table/tbody/tr/td[1]/form/div/center/table/tbody/tr/td[1]/table/tbody/tr[3]/td[2]/p/input"));
			button2.click();
			
			assertEquals( "**Successful Login**",driver.findElement(By.xpath("/html/body/table/tbody/tr/td[1]/big/blockquote/blockquote/font/center/b")).getText());
		
			String loginMessage = driver.findElement(By.xpath("/html/body/table/tbody/tr/td[1]/big/blockquote/blockquote/font/center/b")).getText();
			
		// Writing
		FileInputStream file = new FileInputStream("C:\\Users\\Admin\\Documents\\DemoSiteDDT.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
				XSSFRow row = sheet.getRow(rowNum);
				XSSFCell cell = row.getCell(3);
				if (cell == null) {
					cell = row.createCell(3);
					
				}
				cell.setCellValue(loginMessage);
				XSSFRow row1 = sheet.getRow(rowNum);
					XSSFCell cell1 = row.getCell(4);		
			
				try {	
				
				if (sheet.getRow(rowNum).getCell(2).getStringCellValue().equals(sheet.getRow(rowNum).getCell(3).getStringCellValue())) {
					if (cell1 == null) {
						cell1 = row1.createCell(4);
					}
					cell1.setCellValue("Success");
				}
				}catch (AssertionError e) { if (cell1 == null) {
					cell1 = row1.createCell(4);
					cell1.setCellValue("Failure");
					}
					
				}	
				finally {
				FileOutputStream fileOut = new FileOutputStream("C:\\Users\\admin\\Documents\\DemoSiteDDT.xlsx");
				
				workbook.write(fileOut);
				fileOut.flush();
				fileOut.close();

				file.close();
				}
}
	@After
	public void tearDown() {
		driver.quit();
	}

}
