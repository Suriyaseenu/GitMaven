package org.cts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sample2 {

	public static void main(String[] args) throws IOException {

		File file = new File("F:\\java\\Suriya s\\Sample1\\exceldata\\Maven.xlsx");

		FileInputStream fileInputStream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(fileInputStream);
		Sheet createSheet = workbook.createSheet("daata");
		Row createRow = createSheet.createRow(0);
		Cell createCell = createRow.createCell(2);
		createCell.setCellValue("Maven");
		
		FileOutputStream fileOutputStream = new FileOutputStream(file);
		workbook.write(fileOutputStream);
		System.out.println("done");
		
		
		
		
		
		
		
		
		
		
		
//		Sheet sheet = workbook.getSheet("sheet1");
//		Row row = sheet.getRow(0);
//		Cell cell = row.getCell(3);
//		String stringCellValue = cell.getStringCellValue();
//		System.out.println(stringCellValue);
//		if (stringCellValue.equalsIgnoreCase("DOB")) {
//
//			cell.setCellValue("date of birth");
//		}
//		
//		
//		FileOutputStream f = new FileOutputStream(file);
//		workbook.write(f);
	}

}
