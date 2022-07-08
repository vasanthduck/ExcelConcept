package org.test;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
public static void main(String[] args) throws Throwable {
		
		File f = new File("C:\\Users\\lenovo\\eclipse-workspace\\ExcelConcept\\src\\test\\resources\\ExcelWrite.xlsx");
		Workbook W = new XSSFWorkbook();
		Sheet s = W.createSheet("Excel");
		Row r = s.createRow(0);
		Cell c = r.createCell(0);
		c.setCellValue("bharathi");
		
		FileOutputStream f1 = new FileOutputStream(f);
		W.write(f1);
		
}		
}
