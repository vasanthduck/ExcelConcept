package org.test;

import java.io.File;
import java.io.FileInputStream;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) throws Throwable {
		
		File f = new File("C:\\Users\\lenovo\\eclipse-workspace\\ExcelConcept\\src\\test\\resources\\Book.xlsx");
		FileInputStream f1 = new FileInputStream(f);
		Workbook W = new XSSFWorkbook(f1);
		Sheet s = W.getSheet("Sheet1");
		for(int i=0;i<s.getPhysicalNumberOfRows();i++) {
			Row r =  s.getRow(i);
			for(int j=0;j<r.getPhysicalNumberOfCells();j++) {
				Cell c = r.getCell(j);
				int celltype = c.getCellType();
				if(celltype==1) {
					String value = c.getStringCellValue();
					System.out.println(value);
				}
				else if (celltype==0) {
					if (DateUtil.isCellDateFormatted(c)) {
						Date d = c.getDateCellValue();
						SimpleDateFormat sd = new SimpleDateFormat("dd/mm/yyyy");
						String value = sd.format(d);
						System.out.println(value);
					}
					else {
						double d = c.getNumericCellValue();
						long l = (long)d;
						String value = String.valueOf(l);
						System.out.println(value);
					}
				}
			}
		}
	}
}
