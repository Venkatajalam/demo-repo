package org.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Skills {

	public static void main(String[] args) throws IOException, InterruptedException {
		
		WebDriverManager.chromedriver().setup();
		WebDriver driver=new ChromeDriver();
		
		File file=new File("C:\\Users\\VENKAT\\eclipse-workspace\\Demo1\\Excel\\Ven2.xlsx");
		Workbook workbook=new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Datas");
		
		driver.get("http://demo.automationtesting.in/Register.html");
		WebElement skills = driver.findElement(By.id("Skills"));
		skills.	click();
		Thread.sleep(2000);
		
		Select select =new Select(skills);
		List<WebElement> list = select.getOptions();
		for (int i = 0; i < list.size(); i++) {
			WebElement element = list.get(i);
			String text = element.getText();
			Row Row = sheet.createRow(i);
			Cell cell = Row.createCell(0);
			cell.setCellValue(text);
			FileOutputStream p=new FileOutputStream(file);
			workbook.write(p);
		
		}
		
		
		System.out.println("Done");
		
		
		
		
		
		
		
		
		
		
	}
	
	
	
}
