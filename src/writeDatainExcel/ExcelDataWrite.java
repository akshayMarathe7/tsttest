package writeDatainExcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class ExcelDataWrite 
{

	public static void main(String[] args) throws InterruptedException, IOException 
	{

		System.setProperty("webdriver.chrome.driver", "E:\\Soft\\Selenium\\ChromeDriver1\\chromedriver.exe");
		
		WebDriver driver=new ChromeDriver();
		
		driver.get("http://www.store.demoqa.com");
		
		driver.manage().window().maximize();
		
		Thread.sleep(3000);
		driver.findElement(By.xpath("//div[@id='account']/a")).click();
		
		FileInputStream file= new FileInputStream("E:\\Soft\\Selenium\\Workplace\\ExcelReadAndWrite\\src\\testdata\\Data1.xlsx");
		
		XSSFWorkbook ExcelFile=new XSSFWorkbook(file);
		
		XSSFSheet DesiredSheet = ExcelFile.getSheet("Sheet1");
		
		int TotalRows=DesiredSheet.getLastRowNum();
		System.out.println("TotalRows:- " +TotalRows);
		
		
		for (int i=1;i<=TotalRows;i++)
		{
			Cell cell= DesiredSheet.getRow(i).getCell(0);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			
			
			driver.findElement(By.xpath("//input[@id='log']")).clear();
			driver.findElement(By.xpath("//input[@id='log']")).sendKeys(cell.getStringCellValue());
			
			cell=DesiredSheet.getRow(i).getCell(1);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			
			
			driver.findElement(By.xpath("//input[@id='pwd']")).clear();
			driver.findElement(By.xpath("//input[@id='pwd']")).sendKeys(cell.getStringCellValue());
			
			driver.findElement(By.xpath("//input[@id='login']")).click();
			Thread.sleep(9000);
			System.out.println(driver.getTitle());
//To write Data in file
	
			
			FileOutputStream filO=new FileOutputStream("E:\\Soft\\Selenium\\Workplace\\ExcelReadAndWrite\\src\\testdata\\Data1.xlsx");
			String Passmessage="Pass";
			String FailMessage="Fail";
			
			DesiredSheet.getRow(i).getCell(2);
			
			String Title=driver.getTitle();
			
//			if (Title.equalsIgnoreCase("Your Account | ONLINE STORE"))
			
// Create cell where data needs to be written
//			{
			DesiredSheet.getRow(i).createCell(2).setCellValue(Passmessage);
			ExcelFile.write(filO);
			
//			}
			
//			else
			{
				
//				DesiredSheet.getRow(i).createCell(2).setCellValue(FailMessage);
//				ExcelFile.write(filO);
			}
			
			
			
		}
		
		
		
		
	}

}
