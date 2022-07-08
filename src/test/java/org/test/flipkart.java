package org.test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;

public class flipkart {
	public static void main(String[] args) throws Throwable {
		System.setProperty("webdriver.edge.driver", "C:\\Users\\lenovo\\eclipse-workspace\\UniversityIformation\\Driver\\msedgedriver.exe");
		WebDriver Driver = new EdgeDriver();
		String url = ( "https://www.flipkart.com/");
		Driver.get(url);
		
		Driver.findElement(By.xpath("//button[text()='✕']")).click();
		WebElement search = Driver.findElement(By.name("q"));
		search.sendKeys("redmi mobiles",Keys.ENTER);
		Thread.sleep(4000);
		Driver.findElement(By.xpath("(//div[@class='_4rR01T'])[1]")).click();
		String parent = Driver.getWindowHandle();
		Set<String>child = Driver.getWindowHandles();
		for(String x : child) {
			Driver.switchTo().window(x);
		}
		String text = Driver.findElement(By.xpath("//span[@class='B_NuCI']")).getText();
		System.out.println(text);
		File f = new File("C:\\Users\\lenovo\\eclipse-workspace\\ExcelConcept\\src\\test\\resources\\ExcelWrite.xlsx");
		Workbook W = new XSSFWorkbook();
		Sheet s = W.createSheet("Excel");
		Row r = s.createRow(0);
		Cell c = r.createCell(0);
		c.setCellValue(text);
		
		FileOutputStream f1 = new FileOutputStream(f);
		W.write(f1);
}
}