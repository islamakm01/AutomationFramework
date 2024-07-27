package com.excledata;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ReadData {
	
	@Test
	public void DataRead() throws IOException
	{
		FileInputStream fis = new FileInputStream(System.getProperty("user.dir")+"\\Data\\Book1.xlsx");  //finding the excel sheet
		//Open Excel Sheet
		
		XSSFWorkbook wb = new XSSFWorkbook(fis) ;//open the excel sheet
		XSSFSheet sheet = wb.getSheet("Sheet1");
		
		XSSFRow row = sheet.getRow(5);
		
		int rowdatanumber = sheet.getLastRowNum();
		XSSFCell cell = row.getCell(2);
		
		int colNumber = row.getLastCellNum();
		
		System.out.println(cell);
		
		System.out.println(rowdatanumber);
		
		System.out.println(colNumber);
		
		
		
		System.out.println("=================");
		
		for (int i = 0; i < colNumber; i++) {
			XSSFCell cells = row.getCell(i);
			
			
			System.out.print(cells + " | ");
		}
		
		
	
		
		
	}

}
