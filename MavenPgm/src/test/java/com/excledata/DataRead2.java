package com.excledata;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;


public class DataRead2 {

	@Test
	public void dataRead() throws IOException {
	

	FileInputStream fis = new FileInputStream(System.getProperty("user.dir")+"\\Data\\Book1.xlsx");
	
	XSSFWorkbook wb = new XSSFWorkbook(fis);
	
	XSSFSheet sheet = wb.getSheet("Sheet1");
	
	XSSFRow row = sheet.getRow(4);
	
	int columnNumber = row.getLastCellNum();
	
	for (int i = 0; i < columnNumber; i++) {
		XSSFCell cells = row.getCell(i);
		System.out.print(cells + " | ");
	}
	
	
	
}
	
}

