package readDataFromExcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Exceldata 
{

	@SuppressWarnings("deprecation")
	public static void main(String[] args) throws InterruptedException, Exception 
	{
		
		System.setProperty("webdriver.chrome.driver", "E:\\Soft\\Selenium\\ChromeDriver1\\chromedriver.exe");
		
		WebDriver driver=new ChromeDriver();
		
		driver.get("http://www.store.demoqa.com");
		
		driver.manage().window().maximize();
		
		Thread.sleep(3000);
		
		
		driver.findElement(By.xpath("//div[@id='account']/a")).click();
		
		FileInputStream file = new FileInputStream("E:\\Soft\\Selenium\\Workplace\\ExcelReadAndWrite\\src\\testdata\\Data.xlsx");

		XSSFWorkbook Workbook = new XSSFWorkbook(file);
		
		XSSFSheet desiredSheet = Workbook.getSheet("Sheet1");
//Size of sheet, last row
		
		int lastRow=desiredSheet.getLastRowNum();
		
		for (int i=1;i<=lastRow;i++)
		{
			
			Cell cell = desiredSheet.getRow(i).getCell(0);
			
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(By.xpath("//input[@id='log']")).clear();
			driver.findElement(By.xpath("//input[@id='log']")).sendKeys(cell.getStringCellValue());
			
			cell=desiredSheet.getRow(i).getCell(1);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(By.xpath("//input[@id='pwd']")).clear();
			driver.findElement(By.xpath("//input[@id='pwd']")).sendKeys(cell.getStringCellValue());
			
			driver.findElement(By.xpath("//input[@id='login']")).click();
			Thread.sleep(3000);
			
			
		}
		
		
		
		
		
		
	}

}
