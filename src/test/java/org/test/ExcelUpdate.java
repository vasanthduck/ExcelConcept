package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUpdate {
public static void main(String[] args) throws Throwable {
		
		File f = new File("C:\\Users\\lenovo\\eclipse-workspace\\ExcelConcept\\src\\test\\resources\\ExcelWrite.xlsx");
		FileInputStream f1 = new FileInputStream(f);
		Workbook W = new XSSFWorkbook(f1);
		Sheet s = W.getSheet("Excel");
		Row r = s.getRow(0);
		Cell c = r.getCell(0);
		int celltype = c.getCellType();
		if(celltype==1) {
			String value = c.getStringCellValue();
			if(value.equals("bharathi")) {
				c.setCellValue("Siva");
			}
		}
		FileOutputStream f2 = new FileOutputStream(f);
		W.write(f2);
}
}