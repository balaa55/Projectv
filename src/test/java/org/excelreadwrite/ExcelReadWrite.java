package org.excelreadwrite;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Ignore;
import org.testng.annotations.Test;

public class ExcelReadWrite {
	
	@Test
	public void excelRead() throws IOException {
		File f = new File(System.getProperty("user.dir") + "/src/test/resources/FebAttendance.xlsx");
		FileInputStream input = new FileInputStream(f);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		int totalRows = sheet.getPhysicalNumberOfRows();

		for (int i = 0; i < totalRows; i++) {
			XSSFRow row = sheet.getRow(i); // 0 1 2 3 4 5 6
			int totalcells = row.getPhysicalNumberOfCells();
			for (int j = 0; j < totalcells; j++) {
				XSSFCell cell = row.getCell(j); // 0-0, 0-1, 0-2, 0-3 1-0, 1-1, 1-2.........

				if (cell.getCellType() == CellType.NUMERIC) {
					double numericCellValue = cell.getNumericCellValue();
					System.out.println(numericCellValue + "");
				} else {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue + "");
				}
			}
			System.out.println(" ");
		}
		workbook.close();
	}
	@Ignore
	@Test
	public void excelWrite() throws IOException {
		File f = new File(System.getProperty("user.dir") + "/src/test/resources/FebAttendance.xlsx"); // Opening a file
		FileInputStream input = new FileInputStream(f); // Converting a file to inputstream
		XSSFWorkbook workbook = new XSSFWorkbook(input); // Saving workbook
		XSSFSheet sheet = workbook.getSheet("Sheet1"); // Getting a sheet
		XSSFRow row = sheet.getRow(3);
		
		// Getting a row
		//XSSFSheet newsheet = workbook.createSheet("new");
		
//		editing a cell		
//		XSSFCell cell = row.getCell(3);
//		cell.setCellValue("Y");

//		Creating a new cell and updating
		XSSFCell cell = row.createCell(4); // 
		cell.setCellValue("Informed");

		FileOutputStream out = new FileOutputStream(f);
		workbook.write(out);
		workbook.close();
		out.close();
	}
}



